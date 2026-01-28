Attribute VB_Name = "modExport"
'See modMaterialUpdate for header information

Option Explicit

Public Const cintItemNumberColumn As Integer = 0

Public Const cintPartNumberColumn As Integer = 1

Public Const cintDescriptionColumn As Integer = 2

Public Const cintQuantityColumn As Integer = 3

Public Const cintRoutingNoteMaxLength As Integer = 30

Public strCustomer As String

Public intStandardLot As Integer
    
Public strDefaultLocation As String

Public blnTempWorkBook As Boolean
Public blnTempExcel As Boolean

Public gobjModel As SldWorks.Component2 'temp holding var for recrsive routines

Public gstrPartNumber As String 'temp holding var for recrsive routines, the partnumber it is looking for
Public gboolDefaultConfig As Boolean
Public gstrConfigName As String
Public boolLogError As Boolean
Public boolLogMessage As Boolean
Public strFileName As String
Public gboolStatus As Boolean

'Public gboolTransNU As Boolean    *** Removed 4/8/16 TB

'Public gobjThisFile As ModelDoc2 'the active file we are working with
Public gobjThisFile As Object 'the active file we are working with
Public gintComponents As Integer 'this holds the total number of components'**NEW**
Public Const cstrCADFilePath As String = "M:\"
Public Const cstrPrintFilePath As String = "Q:\"
Public Const cstrFDriveFilePath As String = "F:\"
Public Const cstrOutputFile As String = "I:\Import.prn"
Public Const cstrDocFile As String = "I:\"
Public Const cstrLogFile As String = "I:\Log.txt"
Public Const cstrExcelLookupFile As String = "Laser.xls"


Function QuoteMe(strString) As String

'a little function to make it easier to add quotes to a string, adds a space at the end

    QuoteMe = """" & strString & """ "

End Function

Function AssemblyDepth(strItemNumber As String) As Integer

'returns the depth of the component, just counts periods

' "1" returns 1

' "1.1" returns 2

' "1.1.1 returns 3

    AssemblyDepth = 1

    Dim intPosition As Integer

    For intPosition = 1 To Len(strItemNumber)

        If Mid$(strItemNumber, intPosition, 1) = "." Then

            AssemblyDepth = AssemblyDepth + 1

        End If

    Next intPosition

End Function

 

Sub TraverseComponent(swComp As SldWorks.Component2)

'recursive routine used by GetComponent, digs through the branches of assy structure

    Dim vChildComp                  As Variant

    Dim swChildComp                 As SldWorks.Component2

    Dim swCompConfig                As SldWorks.Configuration

    Dim sPadStr                     As String

    Dim i                           As Long
    
 

    If gobjModel Is Nothing Then

        vChildComp = swComp.GetChildren

        For i = 0 To UBound(vChildComp)

            Set swChildComp = vChildComp(i)
            'Debug.Print "TC:  Looking for swChildComp Path name ' " & swChildComp.GetPathName
            gboolDefaultConfig = False
            gstrConfigName = swChildComp.ReferencedConfiguration
            'If swChildComp.ReferencedConfiguration = "Default" Then
                Debug.Print RemoveInstance(swChildComp.Name)
                If RemoveInstance(swChildComp.Name) = gstrPartNumber Then
                    Set gobjModel = swChildComp
                    If gobjModel Is Nothing Then
                        MsgBox "No objModel Found."
                    End If
                    'Debug.Print "TC: Found Default Child component Part Number = " & gstrPartNumber
                    gboolDefaultConfig = True
                    Exit For
                
                 ElseIf (swChildComp.ReferencedConfiguration) = gstrPartNumber Then
                    Set gobjModel = swChildComp
                    If gobjModel Is Nothing Then
                        MsgBox "No objModel Found."
                    End If
                    'Debug.Print "TC: Found Configuration Child component Part Number = " & swChildComp.GetPathName
                    Exit For
                
                Else

                    TraverseComponent swChildComp

                End If
           
            'Else
                'TraverseComponent swChildComp
            
            'End If

        Next i
    
    End If

End Sub

Function getModelRequested() As Object 'ModelDoc2
Dim TEST As ModelDoc2
'GetComponent()

'gets requested component from current active assembly, used to populate gobjModel

    'Dim gobjModel As IComponent2
    
    

    'Set gobjModel = Nothing

    If gobjThisFile.GetType <> swDocumentTypes_e.swDocPART Then

        Dim swConf                      As SldWorks.Configuration

        Dim swRootComp                  As SldWorks.Component2
        
        Set swConf = gobjThisFile.GetActiveConfiguration
        'Debug.Print "GMR: Seach Assembly Config:                             " & swConf.Name
        
        Set swRootComp = swConf.GetRootComponent

        TraverseComponent swRootComp

        Set getModelRequested = gobjModel.GetModelDoc2

    Else 'working with a part, don't need to go looking for it

        Set getModelRequested = gobjThisFile

    End If

    Set gobjModel = Nothing

End Function

Public Function RemoveInstance(ByVal strName As String) As String

'cleans the file name when recovered from the SW tree

    Dim intPosition As Integer

    intPosition = InStrRev(strName, "/")

    strName = Mid$(strName, intPosition + 1) 'removes the branch, leaves only the file with the instance

    intPosition = InStrRev(strName, "-")

    RemoveInstance = Left$(strName, intPosition - 1) 'removes the instance

End Function

Function GetBOM1(objFile As ModelDoc2) As TableAnnotation

'gets the BOM table

    Set GetBOM = Nothing

    Dim theTableAnnotation As SldWorks.TableAnnotation

    Dim objSelManager As SelectionMgr

    Dim SelObjType As Long

    Dim TableAnnotationType As Long

   

    Set objSelManager = objFile.SelectionManager

    SelObjType = objSelManager.GetSelectedObjectType2(1)

 

    If SelObjType <> swSelANNOTATIONTABLES Then

        MsgBox "You must select a BOM table in the before running."

        Exit Function

    End If

 

    Set theTableAnnotation = objSelManager.GetSelectedObject5(1)

    TableAnnotationType = theTableAnnotation.Type

 

    If theTableAnnotation.Type <> swTableAnnotationType_e.swTableAnnotation_BillOfMaterials Then

        MsgBox "Select a BOM table before running."

        Exit Function

    Else

        Set GetBOM = theTableAnnotation

    End If

    Set theTableAnnotation = Nothing

    Set objSelManager = Nothing

End Function

Private Function GetBOM(swModel As ModelDoc2) As TableAnnotation

    Dim swFeature As Feature

    Dim swBomFeature As BomFeature

    Set swFeature = swModel.FirstFeature

    Set swBomFeature = Nothing

    Set GetBOM = Nothing

   

    Do Until swFeature Is Nothing

        If LCase(swFeature.GetTypeName2) = "tablefolder" Then

            Dim tableFeature As Feature

            Set tableFeature = swFeature.GetFirstSubFeature

           

            Do Until tableFeature Is Nothing

                If LCase(tableFeature.GetTypeName2) = "bomfeat" Then

                    Set swBomFeature = tableFeature.GetSpecificFeature2

                    Set tableFeature = Nothing

                    Exit Do

                End If

                Set tableFeature = tableFeature.GetNextSubFeature

            Loop

            Exit Do

        End If

        Set swFeature = swFeature.GetNextFeature

    Loop

    If Not swBomFeature Is Nothing Then

        Dim vAnnotations As Variant

        vAnnotations = swBomFeature.GetTableAnnotations

        Set GetBOM = vAnnotations(0)

    Else

        MsgBox "Bom table not found.", vbCritical

    End If

    Set swFeature = Nothing

End Function

Sub PopulateParts(objBOMTable As TableAnnotation, ByRef astrPartNames() As String)

'this routine actially fills the array with unique parts

    Dim intRow As Integer

    Dim intCounter As Integer

    Dim strTemp As String

    Dim strDescription As String

    Dim intLineItem As Integer

    For intRow = 1 To objBOMTable.RowCount - 1

        strTemp = Trim$(objBOMTable.Text(intRow, cintPartNumberColumn))
        strDescription = Trim$(objBOMTable.Text(intRow, cintDescriptionColumn))
        'since we find the part by its description, if there is no description, skip out and go to the next line '**NEW**
        If strTemp <> "" Then '**NEW**
            intLineItem = Trim$(AddIfUnique(astrPartNames(), strTemp, strDescription))
        Else
            gintComponents = gintComponents - 1 'no description, lets not count it as a component'**NEW**
        End If '**NEW**
    Next intRow

' --- Test ---
'    intRow = 0
'    While astrPartNames(intRow, 0) <> ""
'        Print #1, "Spot # " & intRow & " in the array is " & astrPartNames(intRow, 0)
'        intRow = intRow + 1
'    Wend

End Sub

Function AddIfUnique(ByRef astrPartNames() As String, strNewItem As String, strDescription As String) As Integer

'adds the item to the array if it isn't already there

'return the array number if added else returns -1

    Dim intIndex As Integer

    Dim objModel As ModelDoc2

    Dim strRevision As String
    
    Dim strPrint As String
    Dim strOutsourceA As String
    Dim OSWC As String
    Dim OSWC_A As String
    Dim strCustNumber As String

    intIndex = 0

    AddIfUnique = -1 'returned if not found

    Do While astrPartNames(intIndex, 0) <> ""
       ' Print #1, "Array Item: " & QuoteMe(astrPartNames(intIndex, 0)) & "NewItem: " & QuoteMe(strNewItem) '--Test
        If astrPartNames(intIndex, 0) = strNewItem Then
        '    Print #1, "Found Duplicate... not adding " & QuoteMe(strNewItem)   '--Test
            GoTo ExitHere

        End If

        intIndex = intIndex + 1

    Loop

    


    'if we got here, then we have a new item, add it

    astrPartNames(intIndex, 0) = strNewItem
    Debug.Print "AddIfUnique: New Item Added :" & strNewItem
    astrPartNames(intIndex, 1) = strDescription

    gstrPartNumber = Trim$(strNewItem)
    'Debug.Print "AddIfUnique: gstrPartNumber = " & gstrPartNumber
    'Debug.Print "AddIfUnique: gstrConfigName = " & gstrConfigName
    Set objModel = getModelRequested
'    Debug.Print objModel.GetPathName
    If gboolDefaultConfig = True Then
        gstrConfigName = ""
    ElseIf gstrConfigName = "Default" Then
        gstrConfigName = ""
    'Else
        'gstrConfigName = gstrPartNumber
    End If
    Debug.Print "AddIfUnique: gstrPartNumber = " & gstrPartNumber
    Debug.Print "objModel: " & objModel.GetTitle
    Debug.Print "AddIfUnique: gstrConfigName = " & gstrConfigName
'    strCustomer = objModel.GetCustomInfoValue(gstrConfigName, "Customer")
    
'    strCustomer = ParseString(strCustomer)
    'Debug.Print gstrPartNumber
    strRevision = objModel.GetCustomInfoValue(gstrConfigName, "Revision")
    
    strRevision = ParseString(strRevision)

    strPrint = objModel.GetCustomInfoValue(gstrConfigName, "Print")
    
    strPrint = ParseString(strPrint)

    astrPartNames(intIndex, 2) = strCustomer

    astrPartNames(intIndex, 3) = strRevision
    
    astrPartNames(intIndex, 4) = strPrint

    OSWC = ""
    OSWC_A = ""

    If objModel.GetCustomInfoValue(gstrConfigName, "OS_WC") <> "" Then
        OSWC = Left(objModel.GetCustomInfoValue(gstrConfigName, "OS_WC"), 3)
    End If
    If objModel.GetCustomInfoValue(gstrConfigName, "OS_WC_A") <> "" Then
        OSWC_A = Left(objModel.GetCustomInfoValue(gstrConfigName, "OS_WC_A"), 3)
    End If

    If objModel.GetCustomInfoValue(gstrConfigName, "ReqOS_A") = "Checked" Then 'Assembly needs OS part created
        'MsgBox gstrPartNumber & " needs Assembly OS"
        astrPartNames(intIndex, 5) = "Assem"
        astrPartNames(intIndex, 6) = OSWC_A & "-" & astrPartNames(intIndex, 0)
        'astrPartNames(intIndex, 6) = objModel.GetCustomInfoValue(gstrConfigName, "OSNumber_A")
    ElseIf objModel.GetCustomInfoValue(gstrConfigName, "rbPartType") = "2" Then 'Regular part needs OS part created
        'MsgBox gstrPartNumber & " needs Part OS"
        astrPartNames(intIndex, 5) = "OSPart"
        astrPartNames(intIndex, 6) = OSWC & "-" & astrPartNames(intIndex, 0)
        'astrPartNames(intIndex, 6) = objModel.GetCustomInfoValue(gstrConfigName, "OSNumber")
    ElseIf objModel.GetCustomInfoValue(gstrConfigName, "rbPartType") = "1" Then 'Regular part needs MP part created
        If objModel.GetCustomInfoValue(gstrConfigName, "rbPartTypeSub") = "0" Then 'Machined Part
            'MsgBox gstrPartNumber & " needs Part MP"
            astrPartNames(intIndex, 5) = "MPPart"
            astrPartNames(intIndex, 6) = "MP-" & astrPartNames(intIndex, 0)
            'astrPartNames(intIndex, 6) = objModel.GetCustomInfoValue(gstrConfigName, "MPNumber")
        ElseIf objModel.GetCustomInfoValue(gstrConfigName, "rbPartTypeSub") = "2" Then 'CUST Supplied Part
            strCustNumber = objModel.GetCustomInfoValue(gstrConfigName, "CustPartNumber")
            If Len(strCustNumber) >= 5 Then
                If UCase(Left(strCustNumber, 5)) <> "CUST-" Then
                    strCustNumber = "CUST-" & strCustNumber
                End If
            End If
            astrPartNames(intIndex, 5) = "CustPart"
            astrPartNames(intIndex, 6) = strCustNumber
            'astrPartNames(intIndex, 6) = objModel.GetCustomInfoValue(gstrConfigName, "CustPartNumber")
        End If
    End If


    AddIfUnique = intIndex

ExitHere:

    Set objModel = Nothing

End Function

Function FileNameWithoutExtension(strFileName As String) As String

'takes a full path and returns just the file name

'ex "c:\desktop\JeffIsCool.txt" returns "JeffIsCool"

    Dim intPosition As Integer

    Dim intPosition2 As Integer

    intPosition = InStrRev(strFileName, "\")

    intPosition2 = InStrRev(strFileName, ".")
    
    If intPosition > 0 Or intPosition2 > 0 Then
        FileNameWithoutExtension = Mid$(strFileName, intPosition + 1, intPosition2 - intPosition - 1)
    Else
        FileNameWithoutExtension = strFileName
    End If

    

End Function

Function IsAssembly(objBOMTable As TableAnnotation, intRow As Integer) As Boolean

'returns True if the line row below is a child of the provided row

    If intRow < objBOMTable.RowCount Then

        Dim intCurrentDepth As Integer

        Dim intNextDepth As Integer
        
        Dim strNextPartNumber As String '**NEW**

        intCurrentDepth = AssemblyDepth(objBOMTable.Text(intRow, cintItemNumberColumn))

        intNextDepth = AssemblyDepth(objBOMTable.Text(intRow + 1, cintItemNumberColumn))

        strNextPartNumber = Trim$(objBOMTable.Text(intRow + 1, cintPartNumberColumn)) '**NEW**
        If intCurrentDepth < intNextDepth And strNextPartNumber <> "" Then '**NEW**

            IsAssembly = True

        Else

            IsAssembly = False

        End If

    Else

        IsAssembly = False

    End If

End Function

Sub PopulateItemMaster(objBOMTable As TableAnnotation)

    Dim astrPartNames() As String 'a list of unique parts found in the BOM table
    
    Dim swApp As SldWorks.SldWorks
    Dim intStdLot As Integer
    Dim swModel As ModelDoc2
    Dim strQuantity As String
    Dim strDesc As String
    Dim strParentName As String
    Dim objModel As ModelDoc2
    Dim MLFlag As Boolean
    Dim strOS_WC As String
    Dim strOS_WC_A As String
    Dim strPrint As String
    Dim strRevision As String

    ReDim astrPartNames(objBOMTable.RowCount - 2, 7)

    PopulateParts objBOMTable, astrPartNames() 'gets array needed for output

    Dim intIndex As Integer
    Dim tempComponents As Integer
    Dim intIndexFix As Integer
    Dim OSWC_A As String
    intIndex = 0
    intIndexFix = 0
    tempComponents = gintComponents
    MLFlag = False

    '*** 1/25/17 TB
    '*** Included IM-ISSUE-SW to allow for Backflushing to be set for Purchased Parts.  If IM-TYPE is being set to 2, then set IM-ISSUE-SW to backflush on Operation  with value of 2
    '***                                                                                If IM-TYPE is being set to 0, then set IM-ISSUE-SW to a value of 0
    

    Print #1, "DECL(IM) ADD IM-KEY IM-DRAWING IM-DESCR IM-REV IM-TYPE IM-CLASS IM-CATALOG IM-COMMODITY IM-SAVE-DEMAND-SW IM-BUYER IM-STOCK-SW IM-PLAN-SW IM-STD-LOT IM-GL-ACCT IM-STD-MAT IM-ISSUE-SW"

    Print #1, "END"

    '*** create IM record for parent part/asm
    
    strParentName = FileNameWithoutExtension(gobjThisFile.GetPathName)
    strParentName = ParseString(strParentName)
    gstrPartNumber = strParentName
    'Set objModel = getModelRequested
    gstrConfigName = ""
    strPrint = gobjThisFile.GetCustomInfoValue(gstrConfigName, "Print")
    strDesc = gobjThisFile.GetCustomInfoValue(gstrConfigName, "Description")
    strRevision = gobjThisFile.GetCustomInfoValue(gstrConfigName, "Revision")
    strPrint = ParseString(strPrint)
    strDesc = ParseString(strDesc)
    
    'create parent  10/13/15
    Print #1, QuoteMe(strParentName) & QuoteMe(strPrint) & QuoteMe(strDesc) & QuoteMe(strRevision) & "1 9 " & QuoteMe(strCustomer) & QuoteMe("F") & "0 " & QuoteMe("2014") & "0 " & "0 " & intStandardLot & " """" 0 0"
    'create parent OS Operation
    OSWC_A = ""
    If gobjThisFile.GetCustomInfoValue(gstrConfigName, "OS_WC_A") <> "" Then
        OSWC_A = Left(gobjThisFile.GetCustomInfoValue(gstrConfigName, "OS_WC_A"), 3)
    End If
    If gobjThisFile.GetCustomInfoValue(gstrConfigName, "ReqOS_A") = "Checked" Then 'Assembly needs OS part created
        'MsgBox gstrPartNumber & " needs Assembly OS"
        Print #1, QuoteMe(OSWC_A & "-" & strParentName) & """"" " & QuoteMe(strDesc) & """""" & " 2 1 " & QuoteMe(strCustomer) & QuoteMe(OSWC_A) & "0 " & QuoteMe("BATCH") & "0 0 " & intStandardLot & " " & QuoteMe("6112.1") & "0 2"
    End If
    
    Do While intIndex <= tempComponents - 1 '**NEW**
        If astrPartNames(intIndex - intIndexFix, 0) <> "" Then
          While Trim$(objBOMTable.Text(intIndex + 1, cintPartNumberColumn)) = ""
            intIndex = intIndex + 1
            intIndexFix = intIndexFix + 1
            tempComponents = tempComponents + 1
          Wend
            gstrPartNumber = astrPartNames(intIndex - intIndexFix, 0)
            Set objModel = getModelRequested
            gstrConfigName = ""
           ' MsgBox objModel.GetCustomInfoValue(gstrConfigName, "Description")
            strQuantity = Trim$(objBOMTable.Text(intIndex + 1, cintQuantityColumn))
            strDesc = astrPartNames(intIndex - intIndexFix, 1)
            strDesc = ParseString(strDesc)
            intStdLot = intStandardLot * strQuantity
            frmProgress.lblTask.Caption = astrPartNames(intIndex - intIndexFix, 0)
            Print #1, QuoteMe(astrPartNames(intIndex - intIndexFix, 0)) & QuoteMe(astrPartNames(intIndex - intIndexFix, 4)) & QuoteMe(strDesc) & QuoteMe(astrPartNames(intIndex - intIndexFix, 3)) & "1 9 " & QuoteMe(astrPartNames(intIndex - intIndexFix, 2)) & QuoteMe("F") & "0 " & QuoteMe("2014") & "0 " & "0 " & intStdLot & " """" 0 0"
            
            '** test to see if OS part needs to be created
            If astrPartNames(intIndex - intIndexFix, 5) = "Assem" Then
                strOS_WC_A = Left(objModel.GetCustomInfoValue(gstrConfigName, "OS_WC_A"), 3)
                Print #1, QuoteMe(astrPartNames(intIndex - intIndexFix, 6)) & """"" " & QuoteMe(strDesc) & """""" & " 2 1 " & QuoteMe(astrPartNames(intIndex - intIndexFix, 2)) & QuoteMe(strOS_WC_A) & "0 " & QuoteMe("BATCH") & "0 0 1 " & QuoteMe("6112.1") & "0 2"
                MLFlag = True
            ElseIf astrPartNames(intIndex - intIndexFix, 5) = "OSPart" Then
                strOS_WC = Left(objModel.GetCustomInfoValue(gstrConfigName, "OS_WC"), 3)
                Print #1, QuoteMe(astrPartNames(intIndex - intIndexFix, 6)) & """"" " & QuoteMe(strDesc) & """""" & " 2 1 " & QuoteMe(astrPartNames(intIndex - intIndexFix, 2)) & QuoteMe(strOS_WC) & "0 " & QuoteMe("BATCH") & "0 0 1 " & QuoteMe("6112.1") & "0 2"
                MLFlag = True
            ElseIf astrPartNames(intIndex - intIndexFix, 5) = "MPPart" Then
                Print #1, QuoteMe(astrPartNames(intIndex - intIndexFix, 6)) & """"" " & QuoteMe(strDesc) & """""" & " 2 1 " & QuoteMe(astrPartNames(intIndex - intIndexFix, 2)) & QuoteMe("MP") & "0 " & QuoteMe("BATCH") & "0 0 1 " & QuoteMe("6110.1") & "0 2"
                MLFlag = True
            ElseIf astrPartNames(intIndex - intIndexFix, 5) = "CustPart" Then
                Print #1, QuoteMe(astrPartNames(intIndex - intIndexFix, 6)) & """"" " & QuoteMe(strDesc) & """""" & " 2 1 " & QuoteMe(astrPartNames(intIndex - intIndexFix, 2)) & QuoteMe("CUS") & "0 " & QuoteMe("BATCH") & "0 0 1 " & QuoteMe("6110.1") & "0 2"
                MLFlag = True 'Is this needed? 4/8/15
            End If
            '**** End test for OS parts
        End If
        intIndex = intIndex + 1
    Loop
    
    'Change made to include creation of ML Locations for ALL parts ## 7/6/2016  TB ##
    MLFlag = True
    
    If MLFlag Then
        
        Print #1, ""
        Print #1, "DECL(ML) ML-IMKEY ML-LOCATION"
        Print #1, "END"
        
        'add in location for top level part
        Print #1, QuoteMe(strParentName) & QuoteMe("PRIMARY")
        
        intIndex = 0
        intIndexFix = 0
        tempComponents = gintComponents
        Do While intIndex <= tempComponents - 1 '**NEW**
            If astrPartNames(intIndex - intIndexFix, 0) <> "" Then
                While Trim$(objBOMTable.Text(intIndex + 1, cintPartNumberColumn)) = ""
                    intIndex = intIndex + 1
                    intIndexFix = intIndexFix + 1
                    tempComponents = tempComponents + 1
                Wend
                Print #1, QuoteMe(astrPartNames(intIndex - intIndexFix, 0)) & QuoteMe("PRIMARY")
                '** test to see if OS/MP part needs to be created
                  'If not OS/MP parts.... now including all parts for shipping program
                If astrPartNames(intIndex - intIndexFix, 5) = "Assem" Or astrPartNames(intIndex - intIndexFix, 5) = "OSPart" Or astrPartNames(intIndex - intIndexFix, 5) = "MPPart" Or astrPartNames(intIndex - intIndexFix, 5) = "CustPart" Then
                    Print #1, QuoteMe(astrPartNames(intIndex - intIndexFix, 6)) & QuoteMe("PRIMARY")
                End If
                  '
                '**** End test for OS/MP parts
            End If
            intIndex = intIndex + 1
        Loop
    End If
   

    Erase astrPartNames

End Sub

Function GetExtension(strFileName As String) As String

'returns all characters after the final .

    Dim intPosition As Integer

    intPosition = InStrRev(strFileName, ".")

    GetExtension = Mid$(strFileName, intPosition + 1)

End Function

Sub PopulateProductStructure(objBOMTable As TableAnnotation)

    Dim astrAssyNames(50) As String 'this array helps keep track of the current assembly branch we are in

    Dim intCurrentDepth As Integer

    Dim intRow As Integer

    astrAssyNames(0) = FileNameWithoutExtension(gobjThisFile.GetPathName)

   

    Dim strPartNumber As String

    Dim strItemNumber As String

    Dim strQuantity As String

    Dim strSubItemNumber As String

 

    Print #1, ""

    Print #1, "DECL(PS) ADD PS-PARENT-KEY PS-SUBORD-KEY PS-REV PS-PIECE-NO PS-QTY-P"

    Print #1, "END"

    For intRow = 1 To objBOMTable.RowCount - 1

        strPartNumber = Trim$(objBOMTable.Text(intRow, cintPartNumberColumn))
        
        If strPartNumber <> "" Then 'Don't add if no part number'**NEW**
        
            frmProgress.lblTask.Caption = strPartNumber

            strItemNumber = Trim$(objBOMTable.Text(intRow, cintItemNumberColumn))

            strQuantity = Trim$(objBOMTable.Text(intRow, cintQuantityColumn))

            strSubItemNumber = GetExtension(strItemNumber)

            intCurrentDepth = AssemblyDepth(strItemNumber)

            If IsAssembly(objBOMTable, intRow) = True Then 'if this is an assembly row remember this

                astrAssyNames(intCurrentDepth) = strPartNumber

            End If
            
            While Len(strSubItemNumber) < 2
                strSubItemNumber = "0" & strSubItemNumber
            Wend
    
            Print #1, QuoteMe(astrAssyNames(intCurrentDepth - 1)) & QuoteMe(strPartNumber) & QuoteMe("COMMON SET") & QuoteMe(strSubItemNumber) & strQuantity
        End If '**NEW**
    Next intRow

   

    Erase astrAssyNames

End Sub

Sub PartMaterialRelationships(objBOMTable As TableAnnotation)

    Dim intRow As Integer
    Dim strPartNumber As String
    Dim strQuantity As String
    Dim strOptiMaterial As String
    Dim strRawWeight As String
    Dim objModel As ModelDoc2
    Dim strERPImport As String
    Dim strCuttingType As String
    Dim strF300Length As String
    Dim strGrainDirection As String
    Dim strWeightLBS As String
    Dim strSheetMetal As String
    Dim strMatl_SF As String
    Dim bRet As Boolean
    Dim strMPNumber As String
    Dim strCustNumber As String
    Dim strPurNumber As String
    Dim strOSNumber As String

    Print #1, ""
    Print #1, "DECL(PS) ADD PS-PARENT-KEY PS-SUBORD-KEY PS-REV PS-PIECE-NO PS-QTY-P PS-DIM-1 PS-ISSUE-SW PS-BFLOCATION-SW PS-BFQTY-SW PS-BFZEROQTY-SW PS-OP-NUM"
    Print #1, "END"

    For intRow = 1 To objBOMTable.RowCount - 1
        strPartNumber = Trim$(objBOMTable.Text(intRow, cintPartNumberColumn))
        If strPartNumber <> "" Then 'no part number, should not be here'**NEW**
            frmProgress.lblTask.Caption = strPartNumber
            strQuantity = Trim$(objBOMTable.Text(intRow, cintQuantityColumn))
            'If IsAssembly(objBOMTable, intRow) = False Then 'if this is an assembly row remember this
                gstrPartNumber = strPartNumber
                Set objModel = getModelRequested
                Debug.Print objModel.GetTitle
                If gboolDefaultConfig = True Then
                    gstrConfigName = ""
                ElseIf gstrConfigName = "Default" Then
                    gstrConfigName = ""
                    'gboolStatus = objModel.ShowConfiguration2(gstrConfigName)
                    'gstrConfigName = gstrPartNumber
                End If
                
            If IsAssembly(objBOMTable, intRow) = False Then 'MOVED **8/25/14
                strERPImport = objModel.GetCustomInfoValue(gstrConfigName, "rbPartType")
                If IsNull(objModel.GetCustomInfoValue(gstrConfigName, "rbPartType")) Or objModel.GetCustomInfoValue(gstrConfigName, "rbPartType") = "" Then
                    strERPImport = 0
                    Print #2, strPartNumber & " :: rbPartType was not set.  Automatically defaulted to 0 (standard part)."
                End If
                If strERPImport = "0" Or strERPImport = "2" Then  'Standard or Outsourced part
                    strOptiMaterial = objModel.GetCustomInfoValue(gstrConfigName, "OptiMaterial")
                    If InStr(1, strOptiMaterial, "#3", vbBinaryCompare) <> 0 Then
                        strGrainDirection = objModel.GetCustomInfoValue(gstrConfigName, "Grain")
                        If strGrainDirection <> "Y" Then
                            Print #2, "GRAIN CHECK: " & strPartNumber & " has #3 in material name, but no grain constraint."
                        End If
                    End If
                    strOptiMaterial = ParseString(strOptiMaterial)
                    'strSheetMetal = objModel.GetCustomInfoValue(gstrConfigName, "swSheetMetal")
                    strSheetMetal = objModel.GetCustomInfoValue(gstrConfigName, "rbMaterialType")
                    'strCuttingType = objModel.GetCustomInfoValue(gstrConfigName, "CuttingType")
'                    strWeightLBS = objModel.Extension.CreateMassProperty.Mass * 2.20462262                             'Commented out 11/14/16 TB
'                    objModel.DeleteCustomInfo2 gstrConfigName, "Weight"                                                'Commented out 11/14/16 TB
'                    objModel.AddCustomInfo3 gstrConfigName, "Weight", swCustomInfoText, Format(strWeightLBS, ".###")   'Commented out 11/14/16 TB
                    If strSheetMetal = "1" Then
                        strRawWeight = "1"
                        strF300Length = objModel.GetCustomInfoValue(gstrConfigName, "F300_Length")
                        If strF300Length = "" Then
                            strF300Length = 0
                        End If
                        '*** modified to backflush on operation 2/21/14 ***
                        Print #1, QuoteMe(strPartNumber) & QuoteMe(strOptiMaterial) & QuoteMe("COMMON SET") & QuoteMe("01") & Format(strRawWeight, ".####") & " " & strF300Length & " 2 1 1 1 20"
                    ElseIf strSheetMetal = "2" Then   'Insulation
                        strMatl_SF = objModel.GetCustomInfoValue(gstrConfigName, "Matl_SF")
                        If strMatl_SF = "" Then
                            strMatl_SF = 0
                        End If
                        Print #1, QuoteMe(strPartNumber) & QuoteMe(strOptiMaterial) & QuoteMe("COMMON SET") & QuoteMe("01") & Format(strMatl_SF, ".####") & " 0" & " 2 1 1 1 20"
                    Else   'strSheetMetal = "0"
                        strRawWeight = objModel.GetCustomInfoValue(gstrConfigName, "RawWeight")
                        bRet = ThicknessCheck.CheckThickness(objModel, strFileName, strOptiMaterial)
                        If bRet = False Then
                            boolLogError = True
                            Print #2, "THICKNESS CHECK: " & strPartNumber & " material doesn't match thickness."
                        End If
                        If strRawWeight = "" Then
                            strRawWeight = "0.0001"
                        End If
                        Print #1, QuoteMe(strPartNumber) & QuoteMe(strOptiMaterial) & QuoteMe("COMMON SET") & QuoteMe("01") & Format(strRawWeight, ".####") & " 0" & " 2 1 1 1 20"
                    End If
                    If strERPImport = "2" Then  'add OS material  '**** ADDED 10/11/2013 ****
                        'strOSNumber = objModel.GetCustomInfoValue(gstrConfigName, "OSNumber")
                        strOSNumber = Left(objModel.GetCustomInfoValue(gstrConfigName, "OS_WC"), 3) & "-" & strPartNumber
                        Print #1, QuoteMe(strPartNumber) & QuoteMe(strOSNumber) & QuoteMe("COMMON SET") & QuoteMe("0") & "1" & " 0" & " 2 1 1 1 20"
                    End If
                ElseIf strERPImport = "1" Then 'Machined part     '**** ADDED 10/11/2013 ****
                    If objModel.GetCustomInfoValue(gstrConfigName, "rbPartTypeSub") = "0" Then 'Machined Part
                        'strMPNumber = objModel.GetCustomInfoValue(gstrConfigName, "MPNumber")
                        strMPNumber = "MP-" & strPartNumber
                        strMPNumber = ParseString(strMPNumber)
                        Print #1, QuoteMe(strPartNumber) & QuoteMe(strMPNumber) & QuoteMe("COMMON SET") & QuoteMe("01") & "1" & " 0" & " 2 1 1 1 20"
                    ElseIf objModel.GetCustomInfoValue(gstrConfigName, "rbPartTypeSub") = "1" Then 'Purchased Part
                        strPurNumber = objModel.GetCustomInfoValue(gstrConfigName, "PurchasedPartNumber")
                        strPurNumber = ParseString(strPurNumber)
                        Print #1, QuoteMe(strPartNumber) & QuoteMe(strPurNumber) & QuoteMe("COMMON SET") & QuoteMe("01") & "1" & " 0" & " 2 1 1 1 10"
                    ElseIf objModel.GetCustomInfoValue(gstrConfigName, "rbPartTypeSub") = "2" Then 'Customer Supplied Part
                        'strCustNumber = objModel.GetCustomInfoValue(gstrConfigName, "CustPartNumber")
                        strCustNumber = objModel.GetCustomInfoValue(gstrConfigName, "CustPartNumber")
                        If Len(strCustNumber) >= 5 Then
                            If UCase(Left(strCustNumber, 5)) <> "CUST-" Then
                                strCustNumber = "CUST-" & strCustNumber
                            End If
                        End If
                        strCustNumber = ParseString(strCustNumber)
                        Print #1, QuoteMe(strPartNumber) & QuoteMe(strCustNumber) & QuoteMe("COMMON SET") & QuoteMe("01") & "1" & " 0" & " 2 1 1 1 10"
                    End If
                End If '**NEW**
            Else 'is an assembly
                If objModel.GetCustomInfoValue(gstrConfigName, "ReqOS_A") = "Checked" Then
                    'MsgBox "Need assembly partMaterialRelationships"
                    strOSNumber = Left(objModel.GetCustomInfoValue(gstrConfigName, "OS_WC_A"), 3) & "-" & strPartNumber
                    'strOSNumber = objModel.GetCustomInfoValue(gstrConfigName, "OSNumber_A")
                    Print #1, QuoteMe(strPartNumber) & QuoteMe(strOSNumber) & QuoteMe("COMMON SET") & QuoteMe("0") & "1" & " 0" & " 2 1 1 1 20"
                End If
                
            End If
            
            Set objModel = Nothing

        End If

    Next intRow

End Sub
Sub CheckRevisitNotes(objBOMTable As TableAnnotation)
    Dim intRow As Integer
    Dim strPartNumber As String
    
    For intRow = 1 To objBOMTable.RowCount - 1
    
        If intRow = objBOMTable.RowCount Then
            strPartNumber = FileNameWithoutExtension(gobjThisFile.GetPathName)
        Else
            strPartNumber = Trim$(objBOMTable.Text(intRow, cintPartNumberColumn))
        End If
        
        If strPartNumber <> "" Then 'strPartNumber populated
            
    
        Else
            Debug.Print "part number blank"
        End If
    Next
    
End Sub

Sub PopulateRouting(objBOMTable As TableAnnotation)

    Dim intRow As Integer
    Dim strPartNumber As String
    Dim strQuantity As String
    Dim strOverrideLocation As String
    Dim strTransfer As String
    Dim strTransferOp As String
    Dim astrSetupRecalc() As String     'used in RecalculateSetupTime 2/23/11
    Dim strPartLocation As String
    Dim strCuttingType As String
    Dim strWorkCenter As String
    Dim strSetup As String
    Dim strRun As String
    Dim strLaserType As String
    Dim objModel As ModelDoc2
    Dim strERPImport As String
    Dim strLocation As String
    Dim strOpNum As String
    Dim strReqOS As String
    Dim strOSNumber_A As String
    Dim strOSLocation_A As String
    Dim strOS_WC_A As String
    Dim strOS_OP_A As String
    Dim bRevisitCB As Boolean
    Dim tbRevisitNote As String
    Dim strOSNumber As String
    Dim strOSLocation As String
    Dim strOS_WC As String
    Dim strOS_OP As String
    Dim strOS_RN As String
    Dim strMPNumber As String
    Dim strPurNumber As String
    Dim strCustNumber As String
    Dim strMPLocation As String
    Dim strMP_RN As String
    Dim boolBlankCheck As Boolean
    Dim boolTransfer As Boolean
    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks
    
    frmProgress.lblProcess.Caption = "Recalculate Setup Time"
    frmProgress.lblTask.Caption = "Starting..."
    
    RecalculateSetupTime objBOMTable, astrSetupRecalc()
    
    frmProgress.lblProcess.Caption = "Populate Routing"
    frmProgress.lblTask.Caption = "Starting..."

    Print #1, ""

    Print #1, "DECL(RT) ADD RT-ITEM-KEY RT-WORKCENTER-KEY  RT-OP-NUM RT-SETUP RT-RUN-STD RT-REV RT-MULT-SEQ"

    Print #1, "END"

    For intRow = 1 To objBOMTable.RowCount - 1

        If intRow = objBOMTable.RowCount Then
            strPartNumber = FileNameWithoutExtension(gobjThisFile.GetPathName)
        Else
            strPartNumber = Trim$(objBOMTable.Text(intRow, cintPartNumberColumn))
        End If
        
        If strPartNumber <> "" Then 'strPartNumber populated
        
            frmProgress.lblTask.Caption = strPartNumber
            strQuantity = Trim$(objBOMTable.Text(intRow, cintQuantityColumn))
            strLocation = 9
        
            If IsAssembly(objBOMTable, intRow) = False Then 'not an assembly

                gstrPartNumber = Trim$(strPartNumber)
                'boolTransfer = False                       *** Removed TRANS 4/7/16  TB
                Set objModel = getModelRequested
                'Debug.Print objModel.GetTitle
                'Debug.Print gstrPartNumber
                If gboolDefaultConfig = True Then
                    gstrConfigName = ""
                ElseIf gstrConfigName = "Default" Then
                    gstrConfigName = ""
                'Else
                '   gstrConfigName = gstrPartNumber
                End If
                
                strERPImport = objModel.GetCustomInfoValue(gstrConfigName, "rbPartType")
                If IsNull(objModel.GetCustomInfoValue(gstrConfigName, "rbPartType")) Or objModel.GetCustomInfoValue(gstrConfigName, "rbPartType") = "" Then
                    strERPImport = 0
                    Print #2, strPartNumber & " :: rbPartType was not set.  Automatically defaulted to 0 (standard part)."
                End If
            
                If strERPImport = "0" Or strERPImport = "2" Then
                        'Begin Calc for strLocation and strTransfer and strTransferOP
                    'strOverrideLocation = objModel.GetCustomInfoValue(gstrConfigName, "PartLocation")
                    'If strOverrideLocation = "1" Then
                        'strDefaultLocation = strLocation
                        'Location = 3 has been decommissioned on 9/15/15
                    '    strPartLocation = objModel.GetCustomInfoValue(gstrConfigName, "Location")
                    '    If strPartLocation = "1" Then 'N1, NU, N2
                    '        strLocation = "F" 'Northern
                    '    ElseIf strPartLocation = "2" = True Then
                    '        strLocation = "N" 'Nu
                    '    ElseIf strPartLocation = "3" = True Then
                    '        strLocation = "D" 'N2
                    '    End If
                    '    If objModel.GetCustomInfoValue(gstrConfigName, "Transfer") <> "1" Then
                    '        boolLogError = True
                    '        Print #2, strPartNumber & "      --> has no Trans operation with Override Location."
                    '    End If
                    'Else
                    '    strLocation = strDefaultLocation
                    'End If

                    'strTransfer = "0"
                    'strTransfer = objModel.GetCustomInfoValue(gstrConfigName, "Transfer")
                    
                   ' If strTransfer = "1" Then
                   '     strTransferOp = "0"
                   '     strTransferOp = objModel.GetCustomInfoValue(gstrConfigName, "TransferOP")
                   ' End If

                'Operation 20

                    '****** MODIFIED 10/11/2013  TB ************
                    '**** This loop seems unnecessary **********
                    '**** Changing loop to look for OP20 value ****
                    
                    If objModel.GetCustomInfoValue(gstrConfigName, "OP20") <> "" Then  'No OP20 wckey
                        strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "OP20")
                        strWorkCenter = Left(strWorkCenter, 4) '4 character WorkCenter
                        strLocation = Left(strWorkCenter, 1) '1 character building code
                        strWorkCenter = Right(strWorkCenter, 3) '3 character WorkCenter
                        strSetup = objModel.GetCustomInfoValue(gstrConfigName, "OP20_S")
                        strRun = objModel.GetCustomInfoValue(gstrConfigName, "OP20_R")
                        'If strLocation = "F" And gboolTransNU Then     *** Removed TRANS 4/7/16  TB
                        '    boolTransfer = True
                        'End If
                        If strRun = "" Or strRun = "." Then
                            strRun = "0"
                            boolLogError = True
                            Print #2, strPartNumber & "      --> has no OP20 run time.  Did you run the cost macro?"
                        End If
                    
                        If strSetup = "." Or strSetup = "" Then
                            strSetup = "0.01"
                            boolLogError = True
                            Print #2, strPartNumber & "      --> has no OP20 setup time. Did you run the cost macro?"
                        End If

                        Select Case strWorkCenter
                            Case Is = 105
                                Print #1, QuoteMe(strPartNumber) & QuoteMe("O" & strWorkCenter) & "10 0 0 " & QuoteMe("COMMON SET") & "0"
                                Print #1, QuoteMe(strPartNumber) & QuoteMe(strLocation & strWorkCenter) & "20 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                            Case Is = 110
                                Print #1, QuoteMe(strPartNumber) & QuoteMe(strLocation & strWorkCenter) & "20 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                            Case Is = 115  'Platino
                                Print #1, QuoteMe(strPartNumber) & QuoteMe("O" & strWorkCenter) & "10 0 0 " & QuoteMe("COMMON SET") & "0"
                                Print #1, QuoteMe(strPartNumber) & QuoteMe(strLocation & strWorkCenter) & "20 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                            Case Is = 120  '5040
                                Print #1, QuoteMe(strPartNumber) & QuoteMe("O" & strWorkCenter) & "10 0 0 " & QuoteMe("COMMON SET") & "0"
                                Print #1, QuoteMe(strPartNumber) & QuoteMe(strLocation & strWorkCenter) & "20 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                            Case Is = 125  '3060
                                Print #1, QuoteMe(strPartNumber) & QuoteMe("O" & strWorkCenter) & "10 0 0 " & QuoteMe("COMMON SET") & "0"
                                Print #1, QuoteMe(strPartNumber) & QuoteMe(strLocation & strWorkCenter) & "20 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                            Case Is = 135  'T6000
                                Print #1, QuoteMe(strPartNumber) & QuoteMe("O" & strWorkCenter) & "10 0 0 " & QuoteMe("COMMON SET") & "0"
                                Print #1, QuoteMe(strPartNumber) & QuoteMe(strLocation & strWorkCenter) & "20 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                            Case Is = 155  'Waterjet
                                Print #1, QuoteMe(strPartNumber) & QuoteMe(strLocation & strWorkCenter) & "20 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                            Case Is = 300  'Saw
                                Print #1, QuoteMe(strPartNumber) & QuoteMe(strLocation & strWorkCenter) & "20 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                            'Case Is = 175  'Python was sold
                                'If strLocation <> "D" Then
                                '    'MsgBox "Python must be at N2. " & strWorkCenter & " location changed on part " & objModel.GetTitle & "."
                                '    strLocation = "D"
                                '    If objModel.GetCustomInfoValue(gstrConfigName, "Transfer") <> "1" Then
                                '        boolLogError = True
                                '        Print #2, strPartNumber & "      -->  has no Trans-Nu after python."
                                '    End If
                                'End If
                                'Print #1, QuoteMe(strPartNumber) & QuoteMe("D" & strWorkCenter) & "20 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                            Case Else
                                Print #1, QuoteMe(strPartNumber) & QuoteMe(strLocation & strWorkCenter) & "20 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                        End Select
                        'If strTransferOp = "20" Then  '10/10/2013  may need to add D location for N2
                        '    Select Case strDefaultLocation
                        '        Case Is = "F"
                        '            strLocation = "F"
                        '        Case Is = "N"
                        '            strLocation = "N"
                        '        Case Is = "D"
                        '            strLocation = "D"
                        '    End Select
                        '    strTransferOp = "21"       '****************** check this  9/18/2015 ************
                        'End If
                    Else
                    'MsgBox "No Op 20 selected on part " & objModel.GetTitle & "."
                        boolLogError = True
                        Print #2, strPartNumber & "      --> has no OP20 workcenter"
                    End If

                'Operation 30

                    If objModel.GetCustomInfoValue(gstrConfigName, "F210") = "1" Then
                        strWorkCenter = "210"
                        strSetup = objModel.GetCustomInfoValue(gstrConfigName, "F210_S")
                        strRun = objModel.GetCustomInfoValue(gstrConfigName, "F210_R")
                        
                        If strSetup = "" Or strSetup = "." Then
                            strSetup = "0"
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has no setup time.  Check that cost macro was run."
                        End If
                        If strRun = "" Or strRun = "." Then
                            strRun = "0"
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has no run time.  Check that cost macro was run."
                        End If
                        
                        If strLocation = "D" Then
                            'MsgBox "Python must be at N2. " & strWorkCenter & " location changed on part " & objModel.GetTitle & "."
                            strLocation = "F"
                            Print #2, "Edge deburring is not at N2. " & strWorkCenter & " location changed to N1 on part " & objModel.GetTitle & "."
                        End If
                        Print #1, QuoteMe(strPartNumber) & QuoteMe(strLocation & strWorkCenter) & "30 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                        
                        'If strTransferOp = "30" Then
                        '    Select Case strDefaultLocation
                        '        Case Is = "F"
                        '            strLocation = "F"
                        '        Case Is = "N"
                        '            strLocation = "N"
                        '        Case Is = "D"
                        '            strLocation = "D"
                        '    End Select
                        '    strTransferOp = "31"
                        'End If
                    End If
                
                'Operation 35
                
                    If objModel.GetCustomInfoValue(gstrConfigName, "F220") = "1" Then
                        strWorkCenter = "220"
                        strSetup = objModel.GetCustomInfoValue(gstrConfigName, "F220_S")
                        strRun = objModel.GetCustomInfoValue(gstrConfigName, "F220_R")
                        
                        If strSetup = "" Or strSetup = "." Then
                            strSetup = "0"
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has no setup time.  Check that cost macro was run."
                        End If
                        If strRun = "" Or strRun = "." Then
                            strRun = "0"
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has no run time.  Check that cost macro was run."
                        End If
                        
                        If strLocation = "D" Then
                            'MsgBox "Python must be at N2. " & strWorkCenter & " location changed on part " & objModel.GetTitle & "."
                            strLocation = "F"
                            Print #2, "Drill cell is not at N2. " & strWorkCenter & " location changed to N1 on part " & objModel.GetTitle & "."
                        End If
                        Print #1, QuoteMe(strPartNumber) & QuoteMe(strLocation & strWorkCenter) & "35 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                        'If (strSetup = "0" Or strRun = "0") Then
                        '    boolLogError = True
                        '    Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        'End If
                        
                        'If strTransferOp = "35" Then
                        '    Select Case strDefaultLocation
                        '        Case Is = "F"
                        '            strLocation = "F"
                        '        Case Is = "N"
                        '            strLocation = "N"
                        '        Case Is = "D"
                        '            strLocation = "D"
                        '    End Select
                        '    strTransferOp = "36"
                        'End If
                    End If
                    
                'Operation 40

                    If objModel.GetCustomInfoValue(gstrConfigName, "PressBrake") = "Checked" Then
                        If strPartNumber = astrSetupRecalc(intRow, 0) Then
                            objModel.DeleteCustomInfo2 "", "F140_S"                                                              'update setup time
                            objModel.AddCustomInfo3 "", "F140_S", swCustomInfoText, Format(astrSetupRecalc(intRow, 3), ".##")    '
                        Else         'should not hit here, but debug in case
                            Debug.Print "Part numbers not equal..."
                            Debug.Print "strPartNumber :: " & strPartNumber & "  intRow :: " & intRow
                            Debug.Print "Array part number :: " & astrSetupRecalc(intRow, 0)
                        End If
                        strWorkCenter = "140"
                        strSetup = objModel.GetCustomInfoValue(gstrConfigName, "F140_S")
                        strRun = objModel.GetCustomInfoValue(gstrConfigName, "F140_R")
                        If strSetup = "" Or strSetup = "." Then
                            strSetup = "0"
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        End If
                        If strRun = "" Or strRun = "." Then
                            strRun = "0"
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        End If
                        
                        If strLocation = "D" Then
                            'MsgBox "Python must be at N2. " & strWorkCenter & " location changed on part " & objModel.GetTitle & "."
                            strLocation = "F"
                            Print #2, "Press brake is not at N2. " & strWorkCenter & " location changed to N1 on part " & objModel.GetTitle & "."
                        End If
                        Print #1, QuoteMe(strPartNumber) & QuoteMe(strLocation & strWorkCenter) & "40 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                        'If (strSetup = "0" Or strRun = "0") Then
                        '    boolLogError = True
                        '    Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        'End If
                        'If strTransferOp = "40" Then
                        '    Select Case strDefaultLocation
                        '        Case Is = "F"
                        '            strLocation = "F"
                        '        Case Is = "N"
                        '            strLocation = "N"
                        '        Case Is = "D"
                        '            strLocation = "D"
                        '    End Select
                        '    strTransferOp = "41"
                        'End If
                    Else
                        If objModel.GetCustomInfoValue(gstrConfigName, "F140_R") <> "" Then
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has press brake run time, but no F140 operation."
                        End If
                        If objModel.GetCustomInfoValue(gstrConfigName, "F140_S") <> "" Then
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has press brake setup time, but no F140 operation."
                        End If
                    End If

                'Operation 50

                    If objModel.GetCustomInfoValue(gstrConfigName, "F325") = "1" Then
                        strWorkCenter = "325"
                        strSetup = objModel.GetCustomInfoValue(gstrConfigName, "F325_S")
                        strRun = objModel.GetCustomInfoValue(gstrConfigName, "F325_R")
                        If strSetup = "" Or strSetup = "." Then
                            strSetup = "0"
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        End If
                        If strRun = "" Or strRun = "." Then
                            strRun = "0"
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        End If
                        Print #1, QuoteMe(strPartNumber) & QuoteMe(strLocation & strWorkCenter) & "50 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                        'If (strSetup = "0" Or strRun = "0") Then
                        '    boolLogError = True
                        '    Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        'End If
                        'If strTransferOp = "50" Then
                        '    Select Case strDefaultLocation
                        '        Case Is = "F"
                        '            strLocation = "F"
                        '        Case Is = "N"
                        '            strLocation = "N"
                        '        Case Is = "D"
                        '            strLocation = "D"
                        '    End Select
                        '    strTransferOp = "51"
                        'End If
                    
                    End If
                    
                'Outsourced Operation   ***ADDED 10/10/2013  TB ****

                    If strERPImport = "2" Then    'Outsourced
                        If objModel.GetCustomInfoValue(gstrConfigName, "OS_WC") = "" Then  'No OS Number input
                            boolLogError = True
                            Print #2, "Outsource Operation not included for part " & strPartNumber
                            strOSNumber = ""
                        Else
                            strOS_WC = Left(objModel.GetCustomInfoValue(gstrConfigName, "OS_WC"), 3) 'Left 3 chars of String
                            strOSNumber = strOS_WC & "-" & strPartNumber
                        End If
                        'Bypassing all OSLocation checks
                        'If objModel.GetCustomInfoValue(gstrConfigName, "OSLocation") = "" Then  'No OS Location input
                        '    boolLogError = True
                        '    Print #2, "OSLocation not included for part " & strPartNumber & ".  Defaulting to N1"
                        '    strOSLocation = "0"
                        'Else
                        '    strOSLocation = objModel.GetCustomInfoValue(gstrConfigName, "OSLocation")
                        'End If
                        strOSLocation = 1
                        
                        If objModel.GetCustomInfoValue(gstrConfigName, "OS_OP") = "" Or objModel.GetCustomInfoValue(gstrConfigName, "OS_OP") = "0" Then   'No Outsource Operation Number input
                            boolLogError = True
                            Print #2, "Outsource Operation Number not included for part " & strPartNumber
                            strOS_OP = ""
                        Else
                            strOS_OP = objModel.GetCustomInfoValue(gstrConfigName, "OS_OP")
                        End If
                        strSetup = 0
                        strRun = 0
                        If strOSLocation = "0" Then
                            strOSLocation = ""
                        ElseIf strOSLocation = "1" Then
                            strOSLocation = "NU"
                        ElseIf strOSLocation = "2" Then
                            strOSLocation = "N2"
                        End If
                        
                        If (strSetup <> "" And strRun <> "") Then
                            Print #1, QuoteMe(strPartNumber) & QuoteMe(strOS_WC & strOSLocation) & strOS_OP & " " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                        End If
                    End If
                
                'Operation Other #1
                    If objModel.GetCustomInfoValue(gstrConfigName, "OtherWC_CB") = "1" Then
                        strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "Other_WC")
                        strSetup = objModel.GetCustomInfoValue(gstrConfigName, "Other_S")
                        strRun = objModel.GetCustomInfoValue(gstrConfigName, "Other_R")
                        strOpNum = objModel.GetCustomInfoValue(gstrConfigName, "OtherOP")
                    
                        'If strTransferOp = strOpNum Then
                        '    Select Case strDefaultLocation
                        '        Case Is = "F"
                        '            strLocation = "F"
                        '        Case Is = "N"
                        '            strLocation = "N"
                        '        Case Is = "D"
                        '            strLocation = "D"
                        '    End Select
                        'End If
                    
                        If strSetup = "" Or strSetup = "." Then
                            strSetup = "0"
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        End If
                        If strRun = "" Or strRun = "." Then
                            strRun = "0"
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        End If
                        Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & strOpNum & " " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                        'If (strSetup = "0" Or strRun = "0") Then
                        '    boolLogError = True
                        '    Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        'End If
                    
                        'If strTransferOp = strOpNum Then
                        '    Select Case strDefaultLocation
                        '        Case Is = "F"
                        '            strLocation = "F"
                        '        Case Is = "N"
                        '            strLocation = "N"
                        '        Case Is = "D"
                        '            strLocation = "D"
                        '    End Select
                        'End If
                    End If
                
                'Operation Other #2
                    If objModel.GetCustomInfoValue(gstrConfigName, "OtherWC_CB2") = "1" Then
                        strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "Other_WC2")
                        strSetup = objModel.GetCustomInfoValue(gstrConfigName, "Other_S2")
                        strRun = objModel.GetCustomInfoValue(gstrConfigName, "Other_R2")
                        strOpNum = objModel.GetCustomInfoValue(gstrConfigName, "Other_OP2")
                    
                        If strSetup = "" Or strSetup = "." Then
                            strSetup = "0"
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        End If
                        If strRun = "" Or strRun = "." Then
                            strRun = "0"
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        End If
                        Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & strOpNum & " " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                        'If (strSetup = "0" Or strRun = "0") Then
                        '    boolLogError = True
                        '    Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        'End If
                
                        'If strTransferOp = strOpNum Then
                        '    Select Case strDefaultLocation
                        '        Case Is = "F"
                        '            strLocation = "F"
                        '        Case Is = "N"
                        '            strLocation = "N"
                        '        Case Is = "D"
                        '            strLocation = "D"
                        '    End Select
                        'End If
                    End If
                
                'Operation Other #3
                    If objModel.GetCustomInfoValue(gstrConfigName, "OtherWC_CB3") = "1" Then
                        strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "Other_WC3")
                        strSetup = objModel.GetCustomInfoValue(gstrConfigName, "Other_S3")
                        strRun = objModel.GetCustomInfoValue(gstrConfigName, "Other_R3")
                        strOpNum = objModel.GetCustomInfoValue(gstrConfigName, "Other_OP3")
                    
                        If strSetup = "" Or strSetup = "." Then
                            strSetup = "0"
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        End If
                        If strRun = "" Or strRun = "." Then
                            strRun = "0"
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        End If
                        Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & strOpNum & " " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                        'If (strSetup = "0" Or strRun = "0") Then
                        '    boolLogError = True
                        '    Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        'End If
                    
                        'If strTransferOp = strOpNum Then
                        '    Select Case strDefaultLocation
                        '        Case Is = "F"
                        '            strLocation = "F"
                        '        Case Is = "N"
                        '            strLocation = "N"
                        '        Case Is = "D"
                        '            strLocation = "D"
                        '    End Select
                        'End If
                    End If
                
                'Operation Other #4
                    If objModel.GetCustomInfoValue(gstrConfigName, "OtherWC_CB4") = "1" Then
                        strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "Other_WC4")
                        strSetup = objModel.GetCustomInfoValue(gstrConfigName, "Other_S4")
                        strRun = objModel.GetCustomInfoValue(gstrConfigName, "Other_R4")
                        strOpNum = objModel.GetCustomInfoValue(gstrConfigName, "Other_OP4")
                    
                        If strSetup = "" Or strSetup = "." Then
                            strSetup = "0"
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        End If
                        If strRun = "" Or strRun = "." Then
                            strRun = "0"
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        End If
                        Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & strOpNum & " " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                        'If (strSetup = "0" Or strRun = "0") Then
                        '    boolLogError = True
                        '    Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        'End If
                    
                        'If strTransferOp = strOpNum Then
                        '    Select Case strDefaultLocation
                        '        Case Is = "F"
                        '            strLocation = "F"
                        '        Case Is = "N"
                        '            strLocation = "N"
                        '        Case Is = "D"
                        '            strLocation = "D"
                        '    End Select
                        'End If
                    End If
                    
                'Operation Other #5
                    If objModel.GetCustomInfoValue(gstrConfigName, "OtherWC_CB5") = "1" Then
                        strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "Other_WC5")
                        strSetup = objModel.GetCustomInfoValue(gstrConfigName, "Other_S5")
                        strRun = objModel.GetCustomInfoValue(gstrConfigName, "Other_R5")
                        strOpNum = objModel.GetCustomInfoValue(gstrConfigName, "Other_OP5")
                    
                        If strSetup = "" Or strSetup = "." Then
                            strSetup = "0"
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        End If
                        If strRun = "" Or strRun = "." Then
                            strRun = "0"
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        End If
                        Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & strOpNum & " " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                        'If (strSetup = "0" Or strRun = "0") Then
                        '    boolLogError = True
                        '    Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        'End If
                    
                        'If strTransferOp = strOpNum Then
                        '    Select Case strDefaultLocation
                        '        Case Is = "F"
                        '            strLocation = "F"
                        '        Case Is = "N"
                        '            strLocation = "N"
                        '        Case Is = "D"
                        '            strLocation = "D"
                        '    End Select
                        'End If
                    End If
                
                'Operation Other #6
                    If objModel.GetCustomInfoValue(gstrConfigName, "OtherWC_CB6") = "1" Then
                        strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "Other_WC6")
                        strSetup = objModel.GetCustomInfoValue(gstrConfigName, "Other_S6")
                        strRun = objModel.GetCustomInfoValue(gstrConfigName, "Other_R6")
                        strOpNum = objModel.GetCustomInfoValue(gstrConfigName, "Other_OP6")
                    
                        If strSetup = "" Or strSetup = "." Then
                            strSetup = "0"
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        End If
                        If strRun = "" Or strRun = "." Then
                            strRun = "0"
                            boolLogError = True
                            Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        End If
                        Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & strOpNum & " " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                        'If (strSetup = "0" Or strRun = "0") Then
                        '    boolLogError = True
                        '    Print #2, strPartNumber & "      -->  has no setup or run time.  Check that cost macro was run."
                        'End If
                    
                        'If strTransferOp = strOpNum Then
                        '    Select Case strDefaultLocation
                        '        Case Is = "F"
                        '            strLocation = "F"
                        '        Case Is = "N"
                        '            strLocation = "N"
                        '        Case Is = "D"
                        '            strLocation = "D"
                        '    End Select
                        'End If
                    End If
                
                    'TRANS OPERATION
                    'If objModel.GetCustomInfoValue(gstrConfigName, "Transfer") = "1" Then
                   ' If boolTransfer = True Then                    *** Removed TRANS 4/7/16  TB
                   '     strTransferOp = 21                            *** Removed TRANS 4/7/16  TB
                        'Select Case strLocation
                        '    Case Is = "F"
                        '        Print #1, QuoteMe(strPartNumber) & QuoteMe("TRANS-N1") & strTransferOp & " " & "0" & " " & "0" & " " & QuoteMe("COMMON SET") & "0"
                        '    Case Is = "N"
                   '             Print #1, QuoteMe(strPartNumber) & QuoteMe("TRANS-NU") & strTransferOp & " " & "0" & " " & "0" & " " & QuoteMe("COMMON SET") & "0"     *** Removed TRANS 4/7/16  TB
                        '    Case Is = "D"
                        '        Print #1, QuoteMe(strPartNumber) & QuoteMe("TRANS-N2") & strTransferOp & " " & "0" & " " & "0" & " " & QuoteMe("COMMON SET") & "0"
                        'End Select
                    ' End If                                        *** Removed TRANS 4/7/16  TB
                ElseIf strERPImport = "1" Then 'Machined part
                    If objModel.GetCustomInfoValue(gstrConfigName, "rbPartTypeSub") = "0" Then 'Machined Part
                    'Machined Part Operation   ***ADDED 10/11/2013  TB ****
                        'If objModel.GetCustomInfoValue(gstrConfigName, "MPNumber") = "" Then  'No MP Number input
                        '    boolLogError = True
                        '    Print #2, "MPNumber not included for part " & strPartNumber
                        '    strMPNumber = ""
                        'Else
                        '    strMPNumber = objModel.GetCustomInfoValue(gstrConfigName, "MPNumber")
                        '    strMPNumber = ParseString(strMPNumber)
                        'End If
                        strMPNumber = "MP-" & strPartNumber
                        strMPNumber = ParseString(strMPNumber)
                        'Bypassing all MPLocation checks
                        'If objModel.GetCustomInfoValue(gstrConfigName, "MPLocation") = "" Then  'No MP Location input
                        '    boolLogError = True
                        '    Print #2, "MPLocation not included for part " & strPartNumber & ".  Defaulting to N1"
                        '    strMPLocation = "0"
                        'Else
                        '    strMPLocation = objModel.GetCustomInfoValue(gstrConfigName, "MPLocation")
                        'End If
                        strMPLocation = 1
                        
                        strSetup = 0
                        strRun = 0
                        If strMPLocation = "0" Then
                            strMPLocation = ""
                        ElseIf strMPLocation = "1" Then
                            strMPLocation = "NU"
                        ElseIf strMPLocation = "2" Then
                            strMPLocation = "N2"
                        End If
                        
                        If (strSetup <> "" And strRun <> "") Then
                            Print #1, QuoteMe(strPartNumber) & QuoteMe("MP" & strMPLocation) & "20" & " " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                        End If
                    ElseIf objModel.GetCustomInfoValue(gstrConfigName, "rbPartTypeSub") = "1" Then 'Purchased part
                        If objModel.GetCustomInfoValue(gstrConfigName, "PurchasedPartNumber") = "" Then  'No Purchased Part Number input
                            boolLogError = True
                            Print #2, "Purchased Part Number not included for part " & strPartNumber
                            strPurNumber = ""
                        Else
                            strPurNumber = objModel.GetCustomInfoValue(gstrConfigName, "PurchasedPartNumber")
                            strPurNumber = ParseString(strPurNumber)
                        End If
                        'Bypassing all MPLocation checks
                        'If objModel.GetCustomInfoValue(gstrConfigName, "MPLocation") = "" Then  'No MP Location input
                        '    boolLogError = True
                        '    Print #2, "Purchased Part Location not included for part " & strPartNumber & ".  Defaulting to N1"
                        '    strMPLocation = "0"
                        'Else
                        '    strMPLocation = objModel.GetCustomInfoValue(gstrConfigName, "MPLocation")
                        'End If
                        strMPLocation = 1
                        
                        strSetup = 0
                        strRun = 0
                        If strMPLocation = "0" Then
                            strMPLocation = ""
                        ElseIf strMPLocation = "1" Then
                            strMPLocation = "N"
                        ElseIf strMPLocation = "2" Then
                            strMPLocation = "N2"
                        End If
                        
                        If (strSetup <> "" And strRun <> "") Then
                            Print #1, QuoteMe(strPartNumber) & QuoteMe(strMPLocation & "PUR") & "10" & " " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                        End If
                    ElseIf objModel.GetCustomInfoValue(gstrConfigName, "rbPartTypeSub") = "2" Then 'Customer Supplied Part
                        If objModel.GetCustomInfoValue(gstrConfigName, "CustPartNumber") = "" Then  'No Cust Number input
                            boolLogError = True
                            Print #2, "Customer Supplied Number not included for part " & strPartNumber
                            strCustNumber = ""
                        Else
                            strCustNumber = objModel.GetCustomInfoValue(gstrConfigName, "CustPartNumber")
                            If Len(strCustNumber) >= 5 Then
                                If UCase(Left(strCustNumber, 5)) <> "CUST-" Then
                                    strCustNumber = "CUST-" & strCustNumber
                                End If
                            End If
                            strCustNumber = ParseString(strCustNumber)
                        End If
                        
                        strSetup = 0
                        strRun = 0
                        If (strSetup <> "" And strRun <> "") Then
                            Print #1, QuoteMe(strPartNumber) & QuoteMe("CUST") & "10" & " " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                        End If
                    End If
                End If  ' End strERPImport   ''' If rbPartType = 0,1,2 Then

            Else  ' IsAssembly = True
                gstrPartNumber = Trim$(strPartNumber)
                'If intRow = objBOMTable.RowCount Then
                'Set objModel = swApp.ActiveDoc
                'Else
                Set objModel = getModelRequested
                'End If
                If gboolDefaultConfig = True Then
                    gstrConfigName = ""
                ElseIf gstrConfigName = "Default" Then
                    gstrConfigName = ""
                'Else
                    'gstrConfigName = gstrPartNumber
                End If
                'Set objModel = GetModelRequested

                '***MODIFIED 10/14/2013 TB new value lbKITPUR ***
                'strLocation = objModel.GetCustomInfoValue(gstrConfigName, "PartLocation")
                'strLocation = objModel.GetCustomInfoValue(gstrConfigName, "lbKITPUR")
                'force op to NKIT
                strLocation = "NKIT"
            
            'OP 10
                If strLocation = "KIT" Or strLocation = "NKIT" Or strLocation = "N2KIT" Then
                    Print #1, QuoteMe(strPartNumber) & QuoteMe(strLocation) & "10 " & "0" & " " & "0" & " " & QuoteMe("COMMON SET") & "0"
                ElseIf strLocation = "PUR" Or strLocation = "NPUR" Or strLocation = "N2PUR" Then
                    Print #1, QuoteMe(strPartNumber) & QuoteMe(strLocation) & "10 " & "0" & " " & "0" & " " & QuoteMe("COMMON SET") & "0"
                End If
            'OP 20
                If objModel.GetCustomInfoValue(gstrConfigName, "OP20") <> "" Then
                    strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "OP20")
                    strSetup = objModel.GetCustomInfoValue(gstrConfigName, "OP20_S")
                    strRun = objModel.GetCustomInfoValue(gstrConfigName, "OP20_R")
                    If strSetup = "" Or strSetup = "." Then
                        strSetup = 0
                    End If
                    If strRun = "" Or strRun = "." Then
                        strRun = 0
                    End If
                    If Not IsNumeric(strSetup) Or Not IsNumeric(strRun) Then
                        MsgBox "Text value invalid...Check log"
                        boolLogError = True
                        Print #2, "Text value where a Numeric value is required :: " & strPartNumber & " :: OP20"
                    End If
                
                    Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & "20 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                End If
            
            'OP 30
                If objModel.GetCustomInfoValue(gstrConfigName, "OP30") <> "" Then
                    strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "OP30")
                    strSetup = objModel.GetCustomInfoValue(gstrConfigName, "OP30_S")
                    strRun = objModel.GetCustomInfoValue(gstrConfigName, "OP30_R")
                    If strSetup = "" Or strSetup = "." Then
                        strSetup = 0
                    End If
                    If strRun = "" Or strRun = "." Then
                        strRun = 0
                    End If
                    If Not IsNumeric(strSetup) Or Not IsNumeric(strRun) Then
                        MsgBox "Text value invalid...Check log"
                        boolLogError = True
                        Print #2, "Text value where a Numeric value is required :: " & strPartNumber & " :: OP30"
                    End If
                    
                    Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & "30 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                End If
            
            'OP 40
                If objModel.GetCustomInfoValue(gstrConfigName, "OP40") <> "" Then
                    strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "OP40")
                    strSetup = objModel.GetCustomInfoValue(gstrConfigName, "OP40_S")
                    strRun = objModel.GetCustomInfoValue(gstrConfigName, "OP40_R")
                    If strSetup = "" Or strSetup = "." Then
                        strSetup = 0
                    End If
                    If strRun = "" Or strRun = "." Then
                        strRun = 0
                    End If
                    If Not IsNumeric(strSetup) Or Not IsNumeric(strRun) Then
                        MsgBox "Text value invalid...Check log"
                        boolLogError = True
                        Print #2, "Text value where a Numeric value is required :: " & strPartNumber & " :: OP40"
                    End If
                    
                    Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & "40 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                End If
            
            'OP 50
                If objModel.GetCustomInfoValue(gstrConfigName, "OP50") <> "" Then
                    strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "OP50")
                    strSetup = objModel.GetCustomInfoValue(gstrConfigName, "OP50_S")
                    strRun = objModel.GetCustomInfoValue(gstrConfigName, "OP50_R")
                    If strSetup = "" Or strSetup = "." Then
                        strSetup = 0
                    End If
                    If strRun = "" Or strRun = "." Then
                        strRun = 0
                    End If
                    If Not IsNumeric(strSetup) Or Not IsNumeric(strRun) Then
                        MsgBox "Text value invalid...Check log"
                        boolLogError = True
                        Print #2, "Text value where a Numeric value is required :: " & strPartNumber & " :: OP50"
                    End If
                    
                    Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & "50 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                End If
            
            'OP 60
                If objModel.GetCustomInfoValue(gstrConfigName, "OP60") <> "" Then
                    strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "OP60")
                    strSetup = objModel.GetCustomInfoValue(gstrConfigName, "OP60_S")
                    strRun = objModel.GetCustomInfoValue(gstrConfigName, "OP60_R")
                    If strSetup = "" Or strSetup = "." Then
                        strSetup = 0
                    End If
                    If strRun = "" Or strRun = "." Then
                        strRun = 0
                    End If
                    If Not IsNumeric(strSetup) Or Not IsNumeric(strRun) Then
                        MsgBox "Text value invalid...Check log"
                        boolLogError = True
                        Print #2, "Text value where a Numeric value is required :: " & strPartNumber & " :: OP60"
                    End If
                    
                    Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & "60 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                End If
            
            'OP 70
                If objModel.GetCustomInfoValue(gstrConfigName, "OP70") <> "" Then
                    strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "OP70")
                    strSetup = objModel.GetCustomInfoValue(gstrConfigName, "OP70_S")
                    strRun = objModel.GetCustomInfoValue(gstrConfigName, "OP70_R")
                    If strSetup = "" Or strSetup = "." Then
                        strSetup = 0
                    End If
                    If strRun = "" Or strRun = "." Then
                        strRun = 0
                    End If
                    If Not IsNumeric(strSetup) Or Not IsNumeric(strRun) Then
                        MsgBox "Text value invalid...Check log"
                        boolLogError = True
                        Print #2, "Text value where a Numeric value is required :: " & strPartNumber & " :: OP70"
                    End If
                    
                    Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & "70 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                End If

            'OP 80
                If objModel.GetCustomInfoValue(gstrConfigName, "OP80") <> "" Then
                    strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "OP80")
                    strSetup = objModel.GetCustomInfoValue(gstrConfigName, "OP80_S")
                    strRun = objModel.GetCustomInfoValue(gstrConfigName, "OP80_R")
                    If strSetup = "" Or strSetup = "." Then
                        strSetup = 0
                    End If
                    If strRun = "" Or strRun = "." Then
                        strRun = 0
                    End If
                    If Not IsNumeric(strSetup) Or Not IsNumeric(strRun) Then
                        MsgBox "Text value invalid...Check log"
                        boolLogError = True
                        Print #2, "Text value where a Numeric value is required :: " & strPartNumber & " :: OP80"
                    End If
                    
                    Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & "80 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                End If
            
            'OP 90
                If objModel.GetCustomInfoValue(gstrConfigName, "OP90") <> "" Then
                    strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "OP90")
                    strSetup = objModel.GetCustomInfoValue(gstrConfigName, "OP90_S")
                    strRun = objModel.GetCustomInfoValue(gstrConfigName, "OP90_R")
                    If strSetup = "" Or strSetup = "." Then
                        strSetup = 0
                    End If
                    If strRun = "" Or strRun = "." Then
                        strRun = 0
                    End If
                    If Not IsNumeric(strSetup) Or Not IsNumeric(strRun) Then
                        MsgBox "Text value invalid...Check log"
                        boolLogError = True
                        Print #2, "Text value where a Numeric value is required :: " & strPartNumber & " :: OP90"
                    End If
                    
                    Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & "90 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                End If
            
            'OP 100
                If objModel.GetCustomInfoValue(gstrConfigName, "OP100") <> "" Then
                    strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "OP100")
                    strSetup = objModel.GetCustomInfoValue(gstrConfigName, "OP100_S")
                    strRun = objModel.GetCustomInfoValue(gstrConfigName, "OP100_R")
                    If strSetup = "" Or strSetup = "." Then
                        strSetup = 0
                    End If
                    If strRun = "" Or strRun = "." Then
                        strRun = 0
                    End If
                    If Not IsNumeric(strSetup) Or Not IsNumeric(strRun) Then
                        MsgBox "Text value invalid...Check log"
                        boolLogError = True
                        Print #2, "Text value where a Numeric value is required :: " & strPartNumber & " :: OP100"
                    End If
                    
                    Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & "100 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                End If
            
            End If
            
            '******* ADDED 10/9/13 TB **************************
            '*** check for RevisitBeforeExport flag is set *****
            '***************************************************
            tbRevisitNote = ""
            bRevisitCB = False
            tbRevisitNote = objModel.GetCustomInfoValue(gstrConfigName, "RevisitNote")
            If objModel.GetCustomInfoValue(gstrConfigName, "RevisitBeforeExport") = "Checked" Then
                bRevisitCB = True
                Print #2, "REVISIT *** " & strPartNumber & ".  Message *** " & tbRevisitNote
                boolLogMessage = True
            End If
            '***************************************************
            
            '****** Outsource assembly route *****
            strReqOS = objModel.GetCustomInfoValue(gstrConfigName, "ReqOS_A")
            If strReqOS = "Checked" Then    'Outsourced
                If objModel.GetCustomInfoValue(gstrConfigName, "OS_WC_A") = "" Then  'No OS Number input
                    boolLogError = True
                    Print #2, "Outsource Operation not included for assembly " & strPartNumber
                    strOSNumber_A = ""
                Else
                    strOS_WC_A = Left(objModel.GetCustomInfoValue(gstrConfigName, "OS_WC_A"), 3) 'Left 3 chars of String
                    strOSNumber_A = strOS_WC_A & "-" & strPartNumber
                End If
                'Bypassing all OSLocation checks
                'If objModel.GetCustomInfoValue(gstrConfigName, "OSLocation_A") = "" Then  'No OS Location input
                '   boolLogError = True
                '    Print #2, "OSLocation not included for assmebly " & strPartNumber & ".  Defaulting to N1"
                '    strOSLocation_A = "0"
                'Else
                '    strOSLocation_A = objModel.GetCustomInfoValue(gstrConfigName, "OSLocation_A")
                'End If
                strOSLocation_A = 1
                
                If objModel.GetCustomInfoValue(gstrConfigName, "OS_OP_A") = "" Or objModel.GetCustomInfoValue(gstrConfigName, "OS_OP_A") = "0" Then   'No Outsource Operation Number input
                    boolLogError = True
                    Print #2, "Outsource Operation Number not included for assembly " & strPartNumber
                    strOS_OP_A = ""
                Else
                    strOS_OP_A = objModel.GetCustomInfoValue(gstrConfigName, "OS_OP_A")
                End If
                strSetup = 0
                strRun = 0
                If strOSLocation_A = "0" Then
                    strOSLocation_A = ""
                ElseIf strOSLocation_A = "1" Then
                    strOSLocation_A = "NU"
                ElseIf strOSLocation_A = "2" Then
                    strOSLocation_A = "N2"
                End If
                    
                If (strSetup <> "" And strRun <> "") Then
                    Print #1, QuoteMe(strPartNumber) & QuoteMe(strOS_WC_A & strOSLocation_A) & strOS_OP_A & " " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                End If
            End If
            
            
            
        Else
            'no part number, should not be here'**NEW**
        End If
        
        Set objModel = Nothing
    Next intRow
    
End Sub
Sub RecalculateSetupTime(objBOMTable As TableAnnotation, ByRef astrSetupRecalc() As String)
    'This routine is used to better calculate press brake setup times for parts using the same material
    Dim intRow As Integer
    Dim intYRow As Integer
    Dim Count As Integer
    Dim maxTime As Double
    Dim newTime As Double
    Dim overrideSetup As Double
    Dim strPartNumber As String
    Dim objModel As ModelDoc2
    
    overrideSetup = frmGlobalValues.txtSetupHours.value
    
    If objBOMTable.RowCount <= 2 Then
        GoTo ExitNow
    Else
        ReDim astrSetupRecalc(objBOMTable.RowCount - 1, 5)
    End If

    For intRow = 1 To objBOMTable.RowCount - 1
        If intRow = objBOMTable.RowCount Then
            strPartNumber = FileNameWithoutExtension(gobjThisFile.GetPathName)
        Else
            strPartNumber = Trim$(objBOMTable.Text(intRow, cintPartNumberColumn))  ' read in part number
        End If
        
        'setup array of parts, material, setup and run times
        If strPartNumber <> "" Then                             'found a part number
            If IsAssembly(objBOMTable, intRow) = False Then
                'put part in array
                gstrPartNumber = Trim$(strPartNumber)           'setting global var for recursive routine
                Set objModel = getModelRequested
                If gboolDefaultConfig = True Then
                    gstrConfigName = ""
                ElseIf gstrConfigName = "Default" Then
                    gstrConfigName = ""
                End If
                
                If objModel.GetCustomInfoValue(gstrConfigName, "PressBrake") = "Checked" Then
                    astrSetupRecalc(intRow, 0) = strPartNumber                                     'part number
                    astrSetupRecalc(intRow, 1) = objModel.GetCustomInfoValue("", "OptiMaterial")   'material
                    astrSetupRecalc(intRow, 2) = objModel.GetCustomInfoValue("", "F140_S_Cost")    'old setup time
                    astrSetupRecalc(intRow, 3) = objModel.GetCustomInfoValue("", "F140_S_Cost")    'new setup time default to old
                    If astrSetupRecalc(intRow, 2) = "" Then                                        'if F140_S_Cost is not set yet
                        astrSetupRecalc(intRow, 2) = objModel.GetCustomInfoValue("", "F140_S")
                        astrSetupRecalc(intRow, 3) = objModel.GetCustomInfoValue("", "F140_S")
                        objModel.DeleteCustomInfo2 "", "F140_S_Cost"                                                                 'default cost to F140_S
                        objModel.AddCustomInfo3 "", "F140_S_Cost", swCustomInfoText, Format(astrSetupRecalc(intRow, 2), ".##")       '2/22/11 Todd
                    End If
                    astrSetupRecalc(intRow, 4) = "False"                                           'has this part been updated yet?
                ElseIf objModel.GetCustomInfoValue(gstrConfigName, "F140_R") <> "" Then
                    'has press brake run time, but no F140 operation
                ElseIf objModel.GetCustomInfoValue(gstrConfigName, "F140_S") <> "" Then
                    'has press brake setup time, but no F140 operation
                End If
            Else
                'assembly part
            End If
        Else
            'no part number
        End If
    Next intRow
    
    For intRow = 1 To objBOMTable.RowCount - 1
        If astrSetupRecalc(intRow, 4) = "False" Then    'has not gone through re-calc yet
            If frmGlobalValues.cbOverride.value = True Then 'override all calculations
                newTime = overrideSetup
            Else                                            'calculate new setup time
                Count = 1     'default Count to 1 for the current part
                If astrSetupRecalc(intRow, 2) <> "" Then 'initial part has been costed
                    maxTime = astrSetupRecalc(intRow, 2)
                    For intYRow = intRow + 1 To objBOMTable.RowCount - 1
                        If astrSetupRecalc(intRow, 1) = astrSetupRecalc(intYRow, 1) Then
                            Count = Count + 1                   'found duplicate material
                            If astrSetupRecalc(intYRow, 2) <> "" Then 'part with same material has been costed
                                If astrSetupRecalc(intYRow, 2) > maxTime Then
                                    maxTime = astrSetupRecalc(intYRow, 2)
                                End If
                            Else 'hasnt been costed
                                MsgBox astrSetupRecalc(intYRow, 0) & " has not been costed.  Please see " & cstrLogFile & " for more details."
                            End If
                        End If
                    Next intYRow
                Else 'hasnt been costed
                    maxTime = 0
                    MsgBox astrSetupRecalc(intRow, 0) & " has not been costed.  Please see " & cstrLogFile & " for more details."
                End If
                'calculate new time
                newTime = ((((maxTime * 60) - cdblBreakSetup) / Count) + cdblBreakSetup) / 60
            End If
            Debug.Print astrSetupRecalc(intRow, 0) & " :: Old Time = " & maxTime & " New Time = " & Round(newTime, 3)
            For intYRow = intRow To objBOMTable.RowCount - 1
                If astrSetupRecalc(intRow, 1) = astrSetupRecalc(intYRow, 1) Then
                    astrSetupRecalc(intYRow, 3) = newTime
                    astrSetupRecalc(intYRow, 4) = "True"
                    Debug.Print "Updated setup time for " & astrSetupRecalc(intYRow, 0)
                End If
            Next intYRow
        End If
    Next intRow
       
ExitNow:
    
End Sub
Sub PopulateParentRoutingNotes(objModel As ModelDoc2)
    Dim intRow As Integer
    Dim strPartNumber As String
    Dim strAuthor As String
    Dim strCADFile As String
    Dim strPrint As String
    Dim strPrintFilePath As String
    Dim strFDate As String
    Dim strExportDate As String
    Dim strF210_RN As String
    Dim strOP20_RN As String
    Dim strOP30_RN As String
    Dim strOP40_RN As String
    Dim strOP50_RN As String
    Dim strOS_OP_A As String
    Dim strOS_RN_A As String
    Dim strTemp As String
    Dim strSend As String
    Dim strRead As String
    Dim outP As String
    Dim Count As Integer
    Dim Length As Integer
    Dim strERPImport As String
    Dim strDrawing As String
    Dim intLineNumber As Integer
    Dim strOtherOP20_RN As String
    Dim boolstatus As Boolean
    Dim longstatus As Integer
    Dim longwarnings As Integer
    Dim strOpNum As String
    Dim strOtherRN As String
    Dim strEdgeClass As String
    Dim strSurfaceFinish As String
    Dim strRevision As String
    Dim strPrint3 As String
    Dim strPrint_temp As String
    Dim strPrintFilePath_temp As String
    Dim strPartNumber_temp As String
    Dim strRevision_temp As String
    Dim strCADFile_temp As String
    Dim strDrawing_temp As String
    Dim AdditionalPrint As String
    Dim strKITPUR As String
    Dim strReqOS As String
    Count = 0

    Print #1, ""
    Print #1, "DECL(RN) ADD RN-ITEM-KEY RN-OP-NUM RN-LINE-NO RN-REV RN-DESCR"
    Print #1, "END"
    
            strPartNumber = objModel.GetTitle
            strPartNumber = FileNameWithoutExtension(strPartNumber)
            'gstrConfigName = ""
             'Operation 10
            If gstrConfigName <> "" Then
                strDrawing = gstrConfigName
            Else
                strDrawing = objModel.GetCustomInfoValue(gstrConfigName, "Drawing")
            End If
            strAuthor = objModel.GetCustomInfoValue(gstrConfigName, "Author")
            strAuthor = "AUTHOR: " & strAuthor
            strAuthor = ParseString(strAuthor)
            strExportDate = Date
            objModel.DeleteCustomInfo2 "", "ExportDate"
            objModel.AddCustomInfo3 "", "ExportDate", swCustomInfoText, strExportDate
            strRevision = ""
            strRevision = objModel.GetCustomInfoValue(gstrConfigName, "Revision")
            strCADFile = cstrCADFilePath & strCustomer & "\"
            objModel.DeleteCustomInfo2 gstrConfigName, "CADFile"
            objModel.AddCustomInfo3 gstrConfigName, "CADFile", swCustomInfoText, strCADFile
            'strDrawing = strDrawing
            strPrintFilePath = cstrPrintFilePath & strCustomer & "\"
            strPrint = objModel.GetCustomInfoValue(gstrConfigName, "Print")
            strPrint = strPrint & ".pdf"
            intLineNumber = 4  'MODIFIED ***10/14/2013 TB  was set to 5
            '**** CHANGES 10/15/2013 TB ***********************************
            '**** ADDED loop to not allow routing notes for PUR routes ****
            '**** ONLY FOR OP 10 for now **********************************
            'strKITPUR = objModel.GetCustomInfoValue(gstrConfigName, "lbKITPUR")
            'force KITPUR to NKIT
            strKITPUR = "NKIT"
            If Not (strKITPUR = "PUR" Or strKITPUR = "NPUR" Or strKITPUR = "N2PUR") Then  'if not a purchase route, then routing notes allowed for OP 10
            
                Print #1, QuoteMe(strPartNumber) & "10 1 " & QuoteMe("COMMON SET") & QuoteMe(strAuthor)
                Print #1, QuoteMe(strPartNumber) & "10 2 " & QuoteMe("COMMON SET") & QuoteMe("DATE: " & strExportDate)
                Print #1, QuoteMe(strPartNumber) & "10 3 " & QuoteMe("COMMON SET") & QuoteMe("------------------------------")
                
                strPartNumber_temp = strPartNumber
                strPrintFilePath_temp = strPrintFilePath
                strPrint_temp = strPrint
                strRevision_temp = strRevision
                
                If objModel.GetCustomInfoValue(gstrConfigName, "AttachPrint") = "1" Then
                
                    Print #1, QuoteMe(strPartNumber) & "10 4 " & QuoteMe("COMMON SET") & QuoteMe("ATTACH PRINT:")
                    intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 10, 5, strPrintFilePath)
                    intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 10, intLineNumber, strPrint)
                    'intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 10, intLineNumber, "------------------------------")
                    'strPartNumber = FileNameWithoutExtension(strPartNumber)  **12/3/15 Removed bc file extension removed already
                    strPartNumber = ParseString(strPartNumber)
                    strPrint = ParseString(strPrint)
                    strPrintFilePath = ParseString(strPrintFilePath)
                    strRevision = ParseString(strRevision)
                    strPrint3 = QuoteMe(strPartNumber) & "," & QuoteMe(strRevision) & "," & QuoteMe(strPrintFilePath & strPrint) & "," & QuoteMe("PRINT") & ",,0,"
                    Print #3, strPrint3
                    '************************************************
                    '****ADDED additional prints 10/14/2013 TB ******
                    '************************************************
                    If objModel.GetCustomInfoValue(gstrConfigName, "cbAttachAddPrint1") = "Checked" Then
                        AdditionalPrint = objModel.GetCustomInfoValue(gstrConfigName, "AdditionalPrint1")
                        AdditionalPrint = AdditionalPrint & ".pdf"
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber_temp, 10, intLineNumber, "------------------------------")
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber_temp, 10, intLineNumber, strPrintFilePath_temp)
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber_temp, 10, intLineNumber, AdditionalPrint)
                        strPrint3 = QuoteMe(strPartNumber) & "," & QuoteMe(strRevision) & "," & QuoteMe(strPrintFilePath & AdditionalPrint) & "," & QuoteMe("PRINT1") & ",,0,"
                        Print #3, strPrint3
                        If objModel.GetCustomInfoValue(gstrConfigName, "cbAttachAddPrint2") = "Checked" Then
                            AdditionalPrint = objModel.GetCustomInfoValue(gstrConfigName, "AdditionalPrint2")
                            AdditionalPrint = AdditionalPrint & ".pdf"
                            intLineNumber = SplitStringIntoPieces("#1", strPartNumber_temp, 10, intLineNumber, "------------------------------")
                            intLineNumber = SplitStringIntoPieces("#1", strPartNumber_temp, 10, intLineNumber, strPrintFilePath_temp)
                            intLineNumber = SplitStringIntoPieces("#1", strPartNumber_temp, 10, intLineNumber, AdditionalPrint)
                            strPrint3 = QuoteMe(strPartNumber) & "," & QuoteMe(strRevision) & "," & QuoteMe(strPrintFilePath & AdditionalPrint) & "," & QuoteMe("PRINT2") & ",,0,"
                            Print #3, strPrint3
                            If objModel.GetCustomInfoValue(gstrConfigName, "cbAttachAddPrint3") = "Checked" Then
                                AdditionalPrint = objModel.GetCustomInfoValue(gstrConfigName, "AdditionalPrint3")
                                AdditionalPrint = AdditionalPrint & ".pdf"
                                intLineNumber = SplitStringIntoPieces("#1", strPartNumber_temp, 10, intLineNumber, "------------------------------")
                                intLineNumber = SplitStringIntoPieces("#1", strPartNumber_temp, 10, intLineNumber, strPrintFilePath_temp)
                                intLineNumber = SplitStringIntoPieces("#1", strPartNumber_temp, 10, intLineNumber, AdditionalPrint)
                                strPrint3 = QuoteMe(strPartNumber) & "," & QuoteMe(strRevision) & "," & QuoteMe(strPrintFilePath & AdditionalPrint) & "," & QuoteMe("PRINT3") & ",,0,"
                                Print #3, strPrint3
                            End If
                        End If
                    End If
                    strPartNumber = strPartNumber_temp
                    strRevision = strRevision_temp
                    intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 10, intLineNumber, "------------------------------")
                    'strDrawing = strDrawing_temp
                    'strCADFile = strCADFile_temp
                End If
            
                If objModel.GetCustomInfoValue(gstrConfigName, "AttachCAD") = "1" Then
                    'If intLineNumber = 5 Then
                    'intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 10, 5, "ATTACH CAD:")
                    'Else
                    'intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 10, intLineNumber, "------------------------------")
                    intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 10, intLineNumber, "ATTACH CAD:")
                    'End If
                    intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 10, intLineNumber, strCADFile)
                    intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 10, intLineNumber, strDrawing)
                    'strDrawing = strDrawing & ".EDRW"
                    strCADFile = ParseString(strCADFile)
                    strDrawing = ParseString(strDrawing)
                    'strPartNumber = FileNameWithoutExtension(strPartNumber)  **12/3/15 Removed bc extension already removed
                    strPartNumber = ParseString(strPartNumber)
                    strRevision = ParseString(strRevision)
                    strPrint3 = QuoteMe(strPartNumber) & "," & QuoteMe(strRevision) & "," & QuoteMe(strCADFile & strDrawing) & "," & QuoteMe("CAD") & ",,0,"
                    Print #3, strPrint3
                    strPartNumber = strPartNumber_temp
                End If
            End If
            
            'Outsource Operation  ****ADDED 8/20/2014 TB****
            strReqOS = objModel.GetCustomInfoValue(gstrConfigName, "ReqOS_A")
            If strReqOS = "Checked" Then  'Outsourced part
                If objModel.GetCustomInfoValue(gstrConfigName, "OS_RN_A") <> "" Then
                    strOS_OP_A = objModel.GetCustomInfoValue(gstrConfigName, "OS_OP_A")
                    strOS_RN_A = objModel.GetCustomInfoValue(gstrConfigName, "OS_RN_A")
                    intLineNumber = SplitStringIntoPieces("#1", strPartNumber, strOS_OP_A, 1, strOS_RN_A)
                End If
            End If
            
             'Operation 20

            If objModel.GetCustomInfoValue(gstrConfigName, "OP20") <> "" Then
                strOP20_RN = objModel.GetCustomInfoValue(gstrConfigName, "OP20_RN")
                intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 20, 1, strOP20_RN)
            End If
            
             'Operation 30

            If objModel.GetCustomInfoValue(gstrConfigName, "OP30") <> "" Then
                strOP30_RN = objModel.GetCustomInfoValue(gstrConfigName, "OP30_RN")
                intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 30, 1, strOP30_RN)
            End If
            
             'Operation 40

            If objModel.GetCustomInfoValue(gstrConfigName, "OP40") <> "" Then
                strOP40_RN = objModel.GetCustomInfoValue(gstrConfigName, "OP40_RN")
                intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 40, 1, strOP40_RN)
            End If
            
            'Operation 50

            If objModel.GetCustomInfoValue(gstrConfigName, "OP50") <> "" Then
                strOP50_RN = objModel.GetCustomInfoValue(gstrConfigName, "OP50_RN")
                intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 50, 1, strOP50_RN)
            End If
            
            'Operation 60
            If objModel.GetCustomInfoValue(gstrConfigName, "OP60") <> "" Then
                strOP50_RN = objModel.GetCustomInfoValue(gstrConfigName, "OP60_RN")
                intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 60, 1, strOP50_RN)
            End If
            
            'Operation 70
            If objModel.GetCustomInfoValue(gstrConfigName, "OP70") <> "" Then
                strOP50_RN = objModel.GetCustomInfoValue(gstrConfigName, "OP70_RN")
                intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 70, 1, strOP50_RN)
            End If
            
            'Operation 80
            If objModel.GetCustomInfoValue(gstrConfigName, "OP80") <> "" Then
                strOP50_RN = objModel.GetCustomInfoValue(gstrConfigName, "OP80_RN")
                intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 80, 1, strOP50_RN)
            End If
            
            'Operation 90
            If objModel.GetCustomInfoValue(gstrConfigName, "OP90") <> "" Then
                strOP50_RN = objModel.GetCustomInfoValue(gstrConfigName, "OP90_RN")
                intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 90, 1, strOP50_RN)
            End If
            
            'Operation 100
            If objModel.GetCustomInfoValue(gstrConfigName, "OP100") <> "" Then
                strOP50_RN = objModel.GetCustomInfoValue(gstrConfigName, "OP100_RN")
                intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 100, 1, strOP50_RN)
            End If

Print #1, " "

End Sub


Sub PopulateRoutingNotes(objBOMTable As TableAnnotation)

    Dim intRow As Integer
    Dim strPartNumber As String
    Dim strAuthor As String
    Dim strCADFile As String
    Dim strPrint As String
    Dim strPrintFilePath As String
    Dim strFDate As String
    Dim strExportDate As String
    Dim strF210_RN As String
    Dim strF220_RN As String
    Dim strOP20_RN As String
    Dim strOP30_RN As String
    Dim strOP40_RN As String
    Dim strOP50_RN As String
    Dim objModel As ModelDoc2
    Dim strRevision As String
    Dim strPrint3 As String
    Dim strOMAXFilePath As String
    Dim strOMAXFile As String
    Dim strPYTHONFilePath As String
    Dim strPYTHONFile As String
    Dim strTUBEFilePath As String
    Dim strTUBEFile As String
    Dim strTemp As String
    Dim strSend As String
    Dim strRead As String
    Dim outP As String
    Dim Count As Integer
    Dim Length As Integer
    Dim strERPImport As String
    Dim strDrawing As String
    Dim DocName As String
    Dim intLineNumber As Integer
    Dim intPosition As Integer
    Dim strOtherOP20_RN As String
    Dim strOP20 As String
    Dim boolstatus As Boolean
    Dim longstatus As Integer
    Dim longwarnings As Integer
    Dim strOpNum As String
    Dim strOtherRN As String
    Dim strEdgeClass As String
    Dim strSurfaceFinish As String
    Dim strOS_RN As String
    Dim strOS_OP As String
    Dim strOS_RN_A As String
    Dim strOS_OP_A As String
    Dim strMP_RN As String
    Dim kFacErrFlag As Boolean
    Dim strPrint_temp As String
    Dim strPrintFilePath_temp As String
    Dim strPartNumber_temp As String
    Dim strRevision_temp As String
    Dim strCADFile_temp As String
    Dim strDrawing_temp As String
    Dim AdditionalPrint As String
    Dim strKITPUR As String
    Dim strReqOS As String

    Count = 0

    Print #1, ""
    Print #1, "DECL(RN) ADD RN-ITEM-KEY RN-OP-NUM RN-LINE-NO RN-REV RN-DESCR"
    Print #1, "END"

    For intRow = 1 To objBOMTable.RowCount - 1

        strPartNumber = Trim$(objBOMTable.Text(intRow, cintPartNumberColumn))
        If strPartNumber <> "" Then 'no part number, should not be here'**NEW**
            Debug.Print strPartNumber
            frmProgress.lblTask.Caption = strPartNumber
            strOP20 = ""
            If IsAssembly(objBOMTable, intRow) = False Then 'if this is an assembly row remember this   'This one is not an assembly

                gstrPartNumber = Trim$(strPartNumber)
                Set objModel = getModelRequested
                If gboolDefaultConfig = True Then
                    gstrConfigName = ""
                ElseIf gstrConfigName = "Default" Then
                    gstrConfigName = ""
                'Else
                    'gstrConfigName = gstrPartNumber
                End If
                'Set objModel = getModelRequested
            
                strERPImport = objModel.GetCustomInfoValue(gstrConfigName, "rbPartType")
                If IsNull(objModel.GetCustomInfoValue(gstrConfigName, "rbPartType")) Or objModel.GetCustomInfoValue(gstrConfigName, "rbPartType") = "" Then
                    strERPImport = 0
                    Print #2, strPartNumber & " :: rbPartType was not set.  Automatically defaulted to 0 (standard part)."
                End If
                
                strOP20 = objModel.GetCustomInfoValue(gstrConfigName, "OP20")
                If strOP20 <> "" Then
                    strOP20 = Left(strOP20, 4)
                    strOP20 = Right(strOP20, 3)
                End If
                'If strERPImport = "0" Or strERPImport = "2" Then  'Std or Outsourced
                    'Operation 20
                    kFacErrFlag = False
                    strAuthor = objModel.GetCustomInfoValue("", "Author")
                    strAuthor = "AUTHOR: " & strAuthor
                    strAuthor = ParseString(strAuthor)
                    strRevision = ""
                    strRevision = objModel.GetCustomInfoValue(gstrConfigName, "Revision")

                    strExportDate = Date
                    objModel.DeleteCustomInfo2 gstrConfigName, "ExportDate"
                    objModel.AddCustomInfo3 gstrConfigName, "ExportDate", swCustomInfoText, strExportDate
                    If gstrConfigName <> "" Then
                        strDrawing = gstrConfigName
                    Else
                        strDrawing = objModel.GetCustomInfoValue(gstrConfigName, "Drawing")
                    End If
                    strDrawing = strDrawing
                    strCADFile = cstrCADFilePath & strCustomer & "\"
                    objModel.DeleteCustomInfo2 gstrConfigName, "CADFile"
                    objModel.AddCustomInfo3 gstrConfigName, "CADFile", swCustomInfoText, strCADFile
               
                    longstatus = 0
                    longwarnings = 0
                    strCADFile = objModel.GetCustomInfoValue(gstrConfigName, "CADFile")
                    'longstatus = objModel.SaveAs3(strCADFile & ".eprt", 0, 0)
                    '*******************************************************************
                    intLineNumber = 4 'default past first 3 hard coded rows
                    Print #1, QuoteMe(strPartNumber) & "20 1 " & QuoteMe("COMMON SET") & QuoteMe(strAuthor)
                    Print #1, QuoteMe(strPartNumber) & "20 2 " & QuoteMe("COMMON SET") & QuoteMe("DATE: " & strExportDate)
                    Print #1, QuoteMe(strPartNumber) & "20 3 " & QuoteMe("COMMON SET") & QuoteMe("------------------------------")
                    strCADFile = ParseString(strCADFile)
                    strRevision = ParseString(strRevision)
                
                    If objModel.GetCustomInfoValue(gstrConfigName, "cbAttachCAD") = "Checked" Then 'print attach Cad and attach to PK
                        Print #1, QuoteMe(strPartNumber) & "20 " & intLineNumber & " " & QuoteMe("COMMON SET") & QuoteMe("ATTACH CAD:")   '**attaches CAD to routing notes***
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 20, intLineNumber + 1, strCADFile)                     '**********************************
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 20, intLineNumber, strDrawing)                         '**********************************
                        'strCADFile = ParseString(strCADFile)
                        strDrawing = ParseString(strDrawing)
                        strPartNumber = ParseString(strPartNumber)
                        'strRevision = ParseString(strRevision)
                        strPrint3 = QuoteMe(strPartNumber) & "," & QuoteMe(strRevision) & "," & QuoteMe(strCADFile & strDrawing) & "," & QuoteMe("CAD") & ",,0,"
                        Print #3, strPrint3  'prints to docattach for pk
                    End If
                    
                    DocName = strDrawing
                    intPosition = InStrRev(DocName, ".")
                    If intPosition = 0 Then
                        DocName = strPartNumber
                    Else
                        DocName = Left$(DocName, intPosition - 1)
                    End If
                    'strDrawing = objModel.GetCustomInfoValue(gstrConfigName, "Model")
                    strPrint3 = QuoteMe(strPartNumber) & "," & QuoteMe(strRevision) & "," & QuoteMe(strCADFile & DocName & ".eprt") & "," & QuoteMe("3D") & ",,0,"
                    Print #3, strPrint3
                
                    strPrintFilePath = cstrPrintFilePath & strCustomer & "\"
                    strPrint = objModel.GetCustomInfoValue(gstrConfigName, "Print")
                    
                    If objModel.GetCustomInfoValue(gstrConfigName, "cbAttachPrint") = "Checked" Then
                        strPrint = ParseString(strPrint)
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 20, intLineNumber, "------------------------------")
                        Print #1, QuoteMe(strPartNumber) & "20 " & intLineNumber & " " & QuoteMe("COMMON SET") & QuoteMe("ATTACH PRINT:")   '**attaches PRINT to routing notes*
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 20, intLineNumber + 1, strPrintFilePath)                 '**********************************
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 20, intLineNumber, strPrint & ".pdf")                             '**********************************
                    End If
                    
                    If strPrint <> "" Then
                        strPrint = strPrint & ".pdf"
                        strPrint = ParseString(strPrint)
                        strPrintFilePath = ParseString(strPrintFilePath)
                        strRevision = ParseString(strRevision)
                        strPrint3 = QuoteMe(strPartNumber) & "," & QuoteMe(strRevision) & "," & QuoteMe(strPrintFilePath & strPrint) & "," & QuoteMe("PRINT") & ",,0,"
                        Print #3, strPrint3
                    End If
                    
                If strERPImport = "0" Or strERPImport = "2" Then  'Std or Outsourced
                    If objModel.GetCustomInfoValue(gstrConfigName, "cbKFactorOK") = "Unchecked" And objModel.GetCustomInfoValue(gstrConfigName, "F325") = "0" Then 'send KFactor Error
                        kFacErrFlag = True
                    End If
                    Call sheetmetal1.CheckBends(objModel, kFacErrFlag)
                    
                    '****** ADDED 3/24/14  By: Todd B ********
                    strOMAXFilePath = cstrFDriveFilePath & "OMAX DXF\" & strCustomer & "\"
                    strTUBEFilePath = cstrFDriveFilePath & "T03\"
                    strTUBEFile = strPartNumber
                    strPYTHONFilePath = cstrFDriveFilePath & "PYTHONX\" & strCustomer & "\"
                    If strOP20 = "110" Then
                        'intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 20, intLineNumber, "------------------------------")
                        'Print #1, QuoteMe(strPartNumber) & "20 " & intLineNumber & " " & QuoteMe("COMMON SET") & QuoteMe("ATTACH TUBE PROGRAM:")   '**attaches LST to routing notes*
                        strTUBEFile = ReplaceDash(strTUBEFile)
                        'intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 20, intLineNumber + 1, strTUBEFilePath)                 '**********************************
                        'intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 20, intLineNumber, strTUBEFile & ".LST")
                        strPrint3 = QuoteMe(strPartNumber) & "," & QuoteMe(strRevision) & "," & QuoteMe(ParseString(strTUBEFilePath & strTUBEFile & ".LST")) & "," & QuoteMe("TUBE") & ",,0,"
                        Print #3, strPrint3
                    ElseIf strOP20 = "155" Then
                        'intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 20, intLineNumber, "------------------------------")
                        'Print #1, QuoteMe(strPartNumber) & "20 " & intLineNumber & " " & QuoteMe("COMMON SET") & QuoteMe("ATTACH OMAX DXF:")   '**attaches DXF to routing notes*
                        strOMAXFile = strPartNumber
                        'intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 20, intLineNumber + 1, strOMAXFilePath)                 '**********************************
                        'intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 20, intLineNumber, strOMAXFile & ".dxf")
                        strPrint3 = QuoteMe(strPartNumber) & "," & QuoteMe(strRevision) & "," & QuoteMe(ParseString(strOMAXFilePath & strOMAXFile & ".dxf")) & "," & QuoteMe("OMAX") & ",,0,"
                        Print #3, strPrint3
                    ElseIf strOP20 = "175" Then '** Added 6/12/14 Todd B ***
                        strPYTHONFile = strPartNumber
                        strPrint3 = QuoteMe(strPartNumber) & "," & QuoteMe(strRevision) & "," & QuoteMe(ParseString(strPYTHONFilePath & strPYTHONFile & ".dxf")) & "," & QuoteMe("PYTHON") & ",,0,"
                        Print #3, strPrint3
                    End If
                    '*************************
                
                    If objModel.GetCustomInfoValue(gstrConfigName, "OP20_RN") <> "" Then
                        strOtherOP20_RN = objModel.GetCustomInfoValue(gstrConfigName, "OP20_RN")
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 20, intLineNumber, "------------------------------")
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 20, intLineNumber, strOtherOP20_RN)
                    End If
                
                    'Operation 30
                    If objModel.GetCustomInfoValue(gstrConfigName, "F210") = "1" Then
                        strF210_RN = objModel.GetCustomInfoValue(gstrConfigName, "F210_RN")
                        strSurfaceFinish = objModel.GetCustomInfoValue(gstrConfigName, "SurfaceFinish")
                        strEdgeClass = objModel.GetCustomInfoValue(gstrConfigName, "EdgeClass")
                    
                        If strSurfaceFinish = "" Then
                            'boolLogError = True
                            'Print #2, strPartNumber & "     --> No F210 Surface Finish"
                        Else
                            strSurfaceFinish = "Surface Finish: " & strSurfaceFinish
                        End If
                    
                        If strEdgeClass = "" Then
                            'boolLogError = True
                            'Print #2, strPartNumber & "     --> No F210 Edge Class"
                        Else
                            strEdgeClass = "Edge Class: " & strEdgeClass
                        End If
                    
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 30, 1, strF210_RN)
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 30, 2, "**************************")
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 30, intLineNumber, strSurfaceFinish)
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 30, intLineNumber, strEdgeClass)
                        
                    End If
                
                    'Operation 35
                    If objModel.GetCustomInfoValue(gstrConfigName, "F220") = "1" Then
                        strF220_RN = objModel.GetCustomInfoValue(gstrConfigName, "F220_RN")
                
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 35, 1, strF220_RN)
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 35, 2, "**************************")
                    End If
                
                    'Operation 40
                    If objModel.GetCustomInfoValue(gstrConfigName, "PressBrake") = "Checked" Then
                        Print #1, QuoteMe(strPartNumber) & "40 1 " & QuoteMe("COMMON SET") & QuoteMe("BRAKE TO CAD")
                    End If
             
                    'Outsource Operation  ****ADDED 10/10/2013 TB****
                    If strERPImport = "2" Then  'Outsourced part
                        If objModel.GetCustomInfoValue(gstrConfigName, "OS_RN") <> "" Then
                            strOS_OP = objModel.GetCustomInfoValue(gstrConfigName, "OS_OP")
                            strOS_RN = objModel.GetCustomInfoValue(gstrConfigName, "OS_RN")
                            intLineNumber = SplitStringIntoPieces("#1", strPartNumber, strOS_OP, 1, strOS_RN)
                        End If
                    End If
            
                    'Operation Other
                    If objModel.GetCustomInfoValue(gstrConfigName, "OtherWC_CB") = "1" Then
                        strOpNum = objModel.GetCustomInfoValue(gstrConfigName, "OtherOP")
                        strOtherRN = objModel.GetCustomInfoValue(gstrConfigName, "Other_RN")
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, strOpNum, 1, strOtherRN)
                    End If
                
                    'Operation Other #2
                    If objModel.GetCustomInfoValue(gstrConfigName, "OtherWC_CB2") = "1" Then
                        strOpNum = objModel.GetCustomInfoValue(gstrConfigName, "Other_OP2")
                        strOtherRN = objModel.GetCustomInfoValue(gstrConfigName, "Other_RN2")
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, strOpNum, 1, strOtherRN)
                    End If
                
                    'Operation Other #3
                    If objModel.GetCustomInfoValue(gstrConfigName, "OtherWC_CB3") = "1" Then
                        strOpNum = objModel.GetCustomInfoValue(gstrConfigName, "Other_OP3")
                        strOtherRN = objModel.GetCustomInfoValue(gstrConfigName, "Other_RN3")
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, strOpNum, 1, strOtherRN)
                    End If
                    
                    'Operation Other #4
                    If objModel.GetCustomInfoValue(gstrConfigName, "OtherWC_CB4") = "1" Then
                        strOpNum = objModel.GetCustomInfoValue(gstrConfigName, "Other_OP4")
                        strOtherRN = objModel.GetCustomInfoValue(gstrConfigName, "Other_RN4")
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, strOpNum, 1, strOtherRN)
                    End If
                    
                    'Operation Other #5
                    If objModel.GetCustomInfoValue(gstrConfigName, "OtherWC_CB5") = "1" Then
                        strOpNum = objModel.GetCustomInfoValue(gstrConfigName, "Other_OP5")
                        strOtherRN = objModel.GetCustomInfoValue(gstrConfigName, "Other_RN5")
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, strOpNum, 1, strOtherRN)
                    End If
                    
                    'Operation Other #6
                    If objModel.GetCustomInfoValue(gstrConfigName, "OtherWC_CB6") = "1" Then
                        strOpNum = objModel.GetCustomInfoValue(gstrConfigName, "Other_OP6")
                        strOtherRN = objModel.GetCustomInfoValue(gstrConfigName, "Other_RN6")
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, strOpNum, 1, strOtherRN)
                    End If
                ElseIf strERPImport = "1" Then 'Machined Part
                    If objModel.GetCustomInfoValue(gstrConfigName, "rbPartTypeSub") = "0" Then 'Machined Part
                        If objModel.GetCustomInfoValue(gstrConfigName, "MP_RN") <> "" Then
                            strMP_RN = objModel.GetCustomInfoValue(gstrConfigName, "MP_RN")
                            intLineNumber = SplitStringIntoPieces("#1", strPartNumber, "20", intLineNumber, strMP_RN)
                        End If
                    End If
                End If   'End If strERPImport = 0,1,2 Then
        
            Else  'Else IsAssembly(objBOMTable, intRow) = False   ' This is an assembly
                gstrPartNumber = Trim$(strPartNumber)
                Set objModel = getModelRequested
                If gboolDefaultConfig = True Then
                    gstrConfigName = ""
                ElseIf gstrConfigName = "Default" Then
                    gstrConfigName = ""
                'Else
                    'gstrConfigName = gstrPartNumber
                End If
            
                strOP20 = objModel.GetCustomInfoValue(gstrConfigName, "OP20")
                If strOP20 <> "" Then
                    strOP20 = Left(strOP20, 4)
                    strOP20 = Right(strOP20, 3)
                End If

                'Operation 10
                If gstrConfigName <> "" Then
                    strDrawing = gstrConfigName
                Else
                    strDrawing = objModel.GetCustomInfoValue(gstrConfigName, "Drawing")
                End If
                strAuthor = objModel.GetCustomInfoValue(gstrConfigName, "Author")
                strAuthor = "AUTHOR: " & strAuthor
                strAuthor = ParseString(strAuthor)
                strRevision = ""
                strRevision = objModel.GetCustomInfoValue(gstrConfigName, "Revision")
                strExportDate = Date
                objModel.DeleteCustomInfo2 "", "ExportDate"
                objModel.AddCustomInfo3 "", "ExportDate", swCustomInfoText, strExportDate
                strCADFile = cstrCADFilePath & strCustomer & "\"
                objModel.DeleteCustomInfo2 gstrConfigName, "CADFile"
                objModel.AddCustomInfo3 gstrConfigName, "CADFile", swCustomInfoText, strCADFile
                strDrawing = strDrawing
                strPrintFilePath = cstrPrintFilePath & strCustomer & "\"
                strPrint = objModel.GetCustomInfoValue(gstrConfigName, "Print")
                strPrint = strPrint & ".pdf"
                intLineNumber = 4
                '**** CHANGES 10/15/2013 TB ***********************************
                '**** ADDED loop to not allow routing notes for PUR routes ****
                '**** ONLY FOR OP 10 for now **********************************
                'strKITPUR = objModel.GetCustomInfoValue(gstrConfigName, "lbKITPUR")
                'force KITPUR to NKIT
                strKITPUR = "NKIT"
                If Not (strKITPUR = "PUR" Or strKITPUR = "NPUR" Or strKITPUR = "N2PUR") Then  'if not a purchase route, then routing notes allowed for OP 10
            
                    Print #1, QuoteMe(strPartNumber) & "10 1 " & QuoteMe("COMMON SET") & QuoteMe(strAuthor)
                    Print #1, QuoteMe(strPartNumber) & "10 2 " & QuoteMe("COMMON SET") & QuoteMe("DATE: " & strExportDate)
                    Print #1, QuoteMe(strPartNumber) & "10 3 " & QuoteMe("COMMON SET") & QuoteMe("------------------------------")
            
                    If objModel.GetCustomInfoValue(gstrConfigName, "AttachPrint") = "1" Then
                        Print #1, QuoteMe(strPartNumber) & "10 4 " & QuoteMe("COMMON SET") & QuoteMe("ATTACH PRINT:")
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 10, intLineNumber + 1, strPrintFilePath)
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 10, intLineNumber, strPrint)
                        'intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 10, intLineNumber, "------------------------------")
                        strPartNumber_temp = strPartNumber
                        strPrintFilePath_temp = strPrintFilePath
                        strPrint_temp = strPrint
                        strRevision_temp = strRevision
                        strPartNumber = ParseString(strPartNumber)
                        strPrintFilePath = ParseString(strPrintFilePath)
                        strPrint = ParseString(strPrint)
                        strRevision = ParseString(strRevision)
                        strPrint3 = QuoteMe(strPartNumber) & "," & QuoteMe(strRevision) & "," & QuoteMe(strPrintFilePath & strPrint) & "," & QuoteMe("PRINT") & ",,0,"
                        Print #3, strPrint3
                        DocName = strPartNumber
                        strCADFile_temp = strCADFile
                        strDrawing_temp = strDrawing
                        strCADFile = ParseString(strCADFile)
                        strDrawing = ParseString(strDrawing)
                        'strPartNumber = ParseString(strPartNumber)
                        'strRevision = ParseString(strRevision)
                        'strDrawing = objModel.GetCustomInfoValue(gstrConfigName, "Model")
                        strPrint3 = QuoteMe(strPartNumber) & "," & QuoteMe(strRevision) & "," & QuoteMe(strCADFile & DocName & ".easm") & "," & QuoteMe("3D") & ",,0,"
                        Print #3, strPrint3
                        '************************************************
                        '****ADDED additional prints 10/14/2013 TB ******
                        '************************************************
                        If objModel.GetCustomInfoValue(gstrConfigName, "cbAttachAddPrint1") = "Checked" Then
                            AdditionalPrint = objModel.GetCustomInfoValue(gstrConfigName, "AdditionalPrint1")
                            AdditionalPrint = AdditionalPrint & ".pdf"
                            intLineNumber = SplitStringIntoPieces("#1", strPartNumber_temp, 10, intLineNumber, "------------------------------")
                            intLineNumber = SplitStringIntoPieces("#1", strPartNumber_temp, 10, intLineNumber, strPrintFilePath_temp)
                            intLineNumber = SplitStringIntoPieces("#1", strPartNumber_temp, 10, intLineNumber, AdditionalPrint)
                            strPrint3 = QuoteMe(strPartNumber) & "," & QuoteMe(strRevision) & "," & QuoteMe(strPrintFilePath & AdditionalPrint) & "," & QuoteMe("PRINT1") & ",,0,"
                            Print #3, strPrint3
                            If objModel.GetCustomInfoValue(gstrConfigName, "cbAttachAddPrint2") = "Checked" Then
                                AdditionalPrint = objModel.GetCustomInfoValue(gstrConfigName, "AdditionalPrint2")
                                AdditionalPrint = AdditionalPrint & ".pdf"
                                intLineNumber = SplitStringIntoPieces("#1", strPartNumber_temp, 10, intLineNumber, "------------------------------")
                                intLineNumber = SplitStringIntoPieces("#1", strPartNumber_temp, 10, intLineNumber, strPrintFilePath_temp)
                                intLineNumber = SplitStringIntoPieces("#1", strPartNumber_temp, 10, intLineNumber, AdditionalPrint)
                                strPrint3 = QuoteMe(strPartNumber) & "," & QuoteMe(strRevision) & "," & QuoteMe(strPrintFilePath & AdditionalPrint) & "," & QuoteMe("PRINT2") & ",,0,"
                                Print #3, strPrint3
                                If objModel.GetCustomInfoValue(gstrConfigName, "cbAttachAddPrint3") = "Checked" Then
                                    AdditionalPrint = objModel.GetCustomInfoValue(gstrConfigName, "AdditionalPrint3")
                                    AdditionalPrint = AdditionalPrint & ".pdf"
                                    intLineNumber = SplitStringIntoPieces("#1", strPartNumber_temp, 10, intLineNumber, "------------------------------")
                                    intLineNumber = SplitStringIntoPieces("#1", strPartNumber_temp, 10, intLineNumber, strPrintFilePath_temp)
                                    intLineNumber = SplitStringIntoPieces("#1", strPartNumber_temp, 10, intLineNumber, AdditionalPrint)
                                    strPrint3 = QuoteMe(strPartNumber) & "," & QuoteMe(strRevision) & "," & QuoteMe(strPrintFilePath & AdditionalPrint) & "," & QuoteMe("PRINT3") & ",,0,"
                                    Print #3, strPrint3
                                End If
                            End If
                        End If
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber_temp, 10, intLineNumber, "------------------------------")
                        strPartNumber = strPartNumber_temp
                        strRevision = strRevision_temp
                        strDrawing = strDrawing_temp
                        strCADFile = strCADFile_temp
                    End If
            
            
                    If objModel.GetCustomInfoValue(gstrConfigName, "AttachCAD") = "1" Then
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 10, intLineNumber, "ATTACH CAD:")
                        
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 10, intLineNumber, strCADFile)
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 10, intLineNumber, strDrawing)
                
                        strRevision = ""
                        strRevision = objModel.GetCustomInfoValue(gstrConfigName, "Revision")
                        If Right(LCase$(strDrawing), 4) <> "edrw" Then
                            strDrawing = strDrawing & ".edrw"
                        End If
                        'strDrawing = strDrawing & ".EDRW"
                        strCADFile = ParseString(strCADFile)
                        strDrawing = ParseString(strDrawing)
                        strPartNumber = ParseString(strPartNumber)
                        strRevision = ParseString(strRevision)
                        strPrint3 = QuoteMe(strPartNumber) & "," & QuoteMe(strRevision) & "," & QuoteMe(strCADFile & strDrawing) & "," & QuoteMe("CAD") & ",,0,"
                        Print #3, strPrint3
                        DocName = strDrawing
                        intPosition = InStrRev(DocName, ".")
                        DocName = Left$(DocName, intPosition - 1)
                        'strDrawing = objModel.GetCustomInfoValue(gstrConfigName, "Model")
                        strPrint3 = QuoteMe(strPartNumber) & "," & QuoteMe(strRevision) & "," & QuoteMe(strCADFile & DocName & ".easm") & "," & QuoteMe("3D") & ",,0,"
                        Print #3, strPrint3
                    End If
                End If  ' End Loop for PUR route check
                'Operation 20
                
                '****** ADDED 3/24/14  By: Todd B ********
                    strOMAXFilePath = cstrFDriveFilePath & "OMAX DXF\" & strCustomer & "\"
                    strPYTHONFilePath = cstrFDriveFilePath & "PYTHONX\" & strCustomer & "\"
                    strTUBEFilePath = cstrFDriveFilePath & "T03\"
                    strTUBEFile = strPartNumber
                    
                    If strOP20 = "110" Then
                        strTUBEFile = ReplaceDash(strTUBEFile)
                        strPrint3 = QuoteMe(strPartNumber) & "," & QuoteMe(strRevision) & "," & QuoteMe(ParseString(strTUBEFilePath & strTUBEFile & ".LST")) & "," & QuoteMe("TUBE") & ",,0,"
                        Print #3, strPrint3
                    ElseIf strOP20 = "155" Then
                        strOMAXFile = strPartNumber
                        strPrint3 = QuoteMe(strPartNumber) & "," & QuoteMe(strRevision) & "," & QuoteMe(ParseString(strOMAXFilePath & strOMAXFile & ".dxf")) & "," & QuoteMe("OMAX") & ",,0,"
                        Print #3, strPrint3
                    ElseIf strOP20 = "175" Then '** Added 6/12/14 Todd B ***
                        strPYTHONFile = strPartNumber
                        strPrint3 = QuoteMe(strPartNumber) & "," & QuoteMe(strRevision) & "," & QuoteMe(ParseString(strPYTHONFilePath & strPYTHONFile & ".dxf")) & "," & QuoteMe("PYTHON") & ",,0,"
                        Print #3, strPrint3
                    End If
                    
                '*************************
                
                If objModel.GetCustomInfoValue(gstrConfigName, "OP20") <> "" Then
                    strOP20_RN = objModel.GetCustomInfoValue(gstrConfigName, "OP20_RN")
                    intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 20, 1, strOP20_RN)
                End If
            
                'Operation 30
                If objModel.GetCustomInfoValue(gstrConfigName, "OP30") <> "" Then
                    strOP30_RN = objModel.GetCustomInfoValue(gstrConfigName, "OP30_RN")
                    intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 30, 1, strOP30_RN)
                End If
            
                'Operation 40
                If objModel.GetCustomInfoValue(gstrConfigName, "OP40") <> "" Then
                    strOP40_RN = objModel.GetCustomInfoValue(gstrConfigName, "OP40_RN")
                    intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 40, 1, strOP40_RN)
                End If
            
                'Operation 50
                If objModel.GetCustomInfoValue(gstrConfigName, "OP50") <> "" Then
                    strOP50_RN = objModel.GetCustomInfoValue(gstrConfigName, "OP50_RN")
                    intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 50, 1, strOP50_RN)
                End If
            
                'Operation 60
                If objModel.GetCustomInfoValue(gstrConfigName, "OP60") <> "" Then
                    strOP50_RN = objModel.GetCustomInfoValue(gstrConfigName, "OP60_RN")
                    intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 60, 1, strOP50_RN)
                End If
            
                'Operation 70
                If objModel.GetCustomInfoValue(gstrConfigName, "OP70") <> "" Then
                    strOP50_RN = objModel.GetCustomInfoValue(gstrConfigName, "OP70_RN")
                    intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 70, 1, strOP50_RN)
                End If
            
                'Operation 80
                If objModel.GetCustomInfoValue(gstrConfigName, "OP80") <> "" Then
                    strOP50_RN = objModel.GetCustomInfoValue(gstrConfigName, "OP80_RN")
                    intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 80, 1, strOP50_RN)
                End If
            
                'Operation 90
                If objModel.GetCustomInfoValue(gstrConfigName, "OP90") <> "" Then
                    strOP50_RN = objModel.GetCustomInfoValue(gstrConfigName, "OP90_RN")
                    intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 90, 1, strOP50_RN)
                End If
            
                'Operation 100
                If objModel.GetCustomInfoValue(gstrConfigName, "OP100") <> "" Then
                    strOP50_RN = objModel.GetCustomInfoValue(gstrConfigName, "OP100_RN")
                    intLineNumber = SplitStringIntoPieces("#1", strPartNumber, 100, 1, strOP50_RN)
                End If
                
                'Outsource Operation  ****ADDED 8/20/14 TB**** copied from part route
                strReqOS = objModel.GetCustomInfoValue(gstrConfigName, "ReqOS_A")
                If strReqOS = "Checked" Then  'Outsourced part
                    If objModel.GetCustomInfoValue(gstrConfigName, "OS_RN_A") <> "" Then
                        strOS_OP_A = objModel.GetCustomInfoValue(gstrConfigName, "OS_OP_A")
                        strOS_RN_A = objModel.GetCustomInfoValue(gstrConfigName, "OS_RN_A")
                        intLineNumber = SplitStringIntoPieces("#1", strPartNumber, strOS_OP_A, 1, strOS_RN_A)
                    End If
                End If
                
                
            End If  'End IsAssembly(objBOMTable, intRow)
            
        Else
            ' strPartNumber = ""  BAD
        End If
        Set objModel = Nothing

    Next intRow

End Sub

Function SplitStringIntoPieces(strFile As String, strPartNumber As String, strOptNum As String, intStartLineNum As Integer, ByVal strRead As String) As Integer

Dim Length As Integer
Dim Count As Integer
Dim strSend As String

Count = intStartLineNum

      Length = Len(strRead)
      While Length > cintRoutingNoteMaxLength
           strSend = Left(strRead, cintRoutingNoteMaxLength)
           strRead = Right(strRead, Length - cintRoutingNoteMaxLength)
           strSend = ParseString(strSend)
           Print #1, QuoteMe(strPartNumber) & strOptNum & " " & Count & " " & QuoteMe("COMMON SET") & QuoteMe(strSend)
           Count = Count + 1
           Length = Len(strRead)
      Wend
      strSend = ParseString(strRead)
      Print #1, QuoteMe(strPartNumber) & strOptNum & " " & Count & " " & QuoteMe("COMMON SET") & QuoteMe(strSend)
        Count = Count + 1
        SplitStringIntoPieces = Count
      'Print #1, QuoteMe(strPartNumber) & strOptNum & " " & Count & " " & QuoteMe("COMMON SET") & QuoteMe("------------------------------")

End Function
Public Function ReplaceDash(str As String) As String

Dim strTemp As String
Dim Count As Integer
Dim Length As Integer
Dim tLength As Integer
Dim X As Integer

    Count = 0
    Length = Len(str)
    strTemp = str
    For X = 1 To Length  'parse through str string
        If Mid(str, X, 1) = Chr(45) Then  'Chr(45) = - (Dash)
            tLength = Len(strTemp)
            'Count = Count + 1  'Add additional for new character to the string
            strTemp = Left(strTemp, Count) & Chr(95) & Right(strTemp, (tLength - Count - 1)) 'swap Chr(45) for  Chr(95) = _ (Underscore)
        End If
        Count = Count + 1   'Move with the X value for temp string
    Next X
    ReplaceDash = strTemp

End Function


Public Function ParseString(str As String) As String

Dim strRoutingNotes As String

Dim Count As Integer

Dim Length As Integer

Dim tLength As Integer

Dim X As Integer

   ' sTest = str

    Count = 0
    
    Length = Len(str)

    strRoutingNotes = str
    
    For X = 1 To Length  'parse through str string

        If Mid(str, X, 1) = Chr(92) Or Mid(str, X, 1) = Chr(34) Then  'Chr(92) = \  Chr (34) = "
            If (X = Length) Or Mid(str, X + 1, 1) <> Chr(92) Then
                tLength = Len(strRoutingNotes)
                Count = Count + 1  'Add additional for new character to the string
                strRoutingNotes = Left(strRoutingNotes, (Count - 1)) & Chr(92) & Right(strRoutingNotes, ((tLength - Count) + 1)) 'add in a \ before character
            ElseIf Mid(str, X + 1, 1) = Chr(92) Then
                X = X + 1
                Count = Count + 1
            End If
            
        End If

        Count = Count + 1   'Move with the X value for temp string

    Next X

    ParseString = strRoutingNotes

    Exit Function

End Function

Sub FixUnits(objBOMTable As TableAnnotation)
    Dim intRow As Integer

    Dim strPartNumber As String

    Dim objModel As ModelDoc2
    
    Dim boolRef As Boolean
    
    'Dim objModelExt As ModelDocExtension
    
    Dim CurUnits As Integer
    
    For intRow = 1 To objBOMTable.RowCount - 1

        strPartNumber = Trim$(objBOMTable.Text(intRow, cintPartNumberColumn))
        
        If strPartNumber <> "" Then 'Don't add if no part number'**NEW**
            gstrPartNumber = Trim$(strPartNumber)
    
            Set objModel = getModelRequested
            If gboolDefaultConfig = True Then
                gstrConfigName = ""
            ElseIf gstrConfigName = "Default" Then
                gstrConfigName = ""
            'Else
            '    gstrConfigName = gstrPartNumber
            End If
            CurUnits = 0
            CurUnits = objModel.GetUserPreferenceIntegerValue(swUnitsMassPropMass)
            CurUnits = objModel.Extension.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsMassPropMass, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
            CurUnits = objModel.Extension.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
            If CurUnits <> swUnitSystem_IPS Then
                'objModel.SetUserPreferenceIntegerValue swUnitsMassPropMass, swUnitsMassPropMass_Pounds
                boolRef = objModel.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, swUserPreferenceOption_e.swDetailingNoOptionSpecified, swUnitSystem_e.swUnitSystem_IPS)
                CurUnits = objModel.Extension.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
                boolRef = objModel.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsMassPropMass, swUserPreferenceOption_e.swDetailingNoOptionSpecified, swUnitsMassPropMass_e.swUnitsMassPropMass_Pounds)
                CurUnits = objModel.Extension.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsMassPropMass, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
            End If
            
            CurUnits = objModel.Extension.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsMassPropMass, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
            If CurUnits <> swUnitsMassPropMass_e.swUnitsMassPropMass_Pounds Then
                'objModel.SetUserPreferenceIntegerValue swUnitsMassPropMass, swUnitsMassPropMass_Pounds
                boolRef = objModel.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, swUserPreferenceOption_e.swDetailingNoOptionSpecified, swUnitSystem_e.swUnitSystem_IPS)
                CurUnits = objModel.Extension.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitSystem, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
                boolRef = objModel.Extension.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsMassPropMass, swUserPreferenceOption_e.swDetailingNoOptionSpecified, swUnitsMassPropMass_e.swUnitsMassPropMass_Pounds)
                CurUnits = objModel.Extension.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsMassPropMass, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
            End If
            CurUnits = objModel.GetUserPreferenceIntegerValue(swUnitsMassPropMass)
            CurUnits = objModel.Extension.GetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swUnitsMassPropMass, swUserPreferenceOption_e.swDetailingNoOptionSpecified)
            
        End If
        Set objModel = Nothing

    Next intRow

End Sub

Public Sub ExportBOM()
    blnTempExcel = False
    blnTempWorkBook = False
    
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swFeat As SldWorks.Feature
    Dim swTableFeat As SldWorks.Feature
    Dim objBOMTable As SldWorks.TableAnnotation
    Dim Customer As String
    
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set objBOMTable = swModel.Extension.InsertBomTable3("C:\Program Files\SolidWorks Corp\SolidWorks\lang\english\bom-standard.sldbomtbt", 0, 0, swBomType_e.swBomType_Indented, "Default", False, swNumberingType_Detailed, False)
    
    strFileName = modMaterialUpdate.ReturnExcelFile(swApp.GetCurrentMacroPathName)

    Customer = swModel.GetCustomInfoValue("", "Customer")
    
    frmGlobalValues.txtCatalogCode = Customer

    If Not objBOMTable Is Nothing Then
        frmGlobalValues.Show
        'frmGlobalValues.ZOrder 0
    End If
    
    swModel.EditRebuild3
    
    Set swFeat = swModel.FirstFeature
    
    While Not swFeat Is Nothing
        Debug.Print swFeat.GetTypeName

        If "TableFolder" = swFeat.GetTypeName Then
        
            Set swTableFeat = swFeat.GetFirstSubFeature

            While Not swTableFeat Is Nothing
            
                swTableFeat.Select2 True, 0
                Set swTableFeat = swTableFeat.GetNextSubFeature
                
            Wend

        End If

        Set swFeat = swFeat.GetNextFeature

    Wend

    swModel.Extension.DeleteSelection2 swDeleteSelectionOptions_e.swDelete_Absorbed
   
    
End Sub

Sub SaveAsEDrawing(objModel As ModelDoc2, strCustomer As String)
Dim boolstatus As Boolean
Dim longstatus As Integer
Dim longwarnings As Integer
Dim filePath As String
Dim NewFilePath As String
Dim strCADFile As String
Dim DocName As String
Dim modelType As Long

modelType = objModel.GetType
longstatus = 0
longwarnings = 0
boolLogError = False
DocName = objModel.GetTitle
DocName = FileNameWithoutExtension(DocName)
If modelType = SwConst.swDocASSEMBLY Then
    longstatus = objModel.SaveAs3(cstrCADFilePath & strCustomer & "\" & DocName & ".easm", 0, 0)
    If longstatus = 1 Then
        frmResults.lblEdrawing.Caption = "FAILED!"
    ElseIf longstatus = 0 Then
        frmResults.lblEdrawing.Caption = "Successful"
    End If
Else
    frmResults.lblEdrawing.Caption = "Not an Assembly."
End If

End Sub


Sub PopulateParentRoute(objModel As ModelDoc2)

    Dim strPartNumber As String
    Dim strQuantity As String
    Dim strOverrideLocation As String
    Dim strPartLocation As String
    Dim strWorkCenter As String
    Dim strSetup As String
    Dim strRun As String
    Dim strERPImport As String
    Dim strLocation As String
    Dim strOpNum As String
    Dim strReqOS As String
    Dim strOSNumber_A As String
    Dim strOSLocation_A As String
    Dim strOS_WC_A As String
    Dim strOS_OP_A As String
    Dim swConfig As SldWorks.Configuration
    Set swConfig = objModel.GetActiveConfiguration
    
    Dim tbRevisitNote As String
    Dim bRevisitCB As Boolean
   

    Print #1, ""

    Print #1, "DECL(RT) ADD RT-ITEM-KEY RT-WORKCENTER-KEY  RT-OP-NUM RT-SETUP RT-RUN-STD RT-REV RT-MULT-SEQ"

    Print #1, "END"

            strPartNumber = objModel.GetTitle
            strPartNumber = FileNameWithoutExtension(strPartNumber)
            'strOverrideLocation = objModel.GetCustomInfoValue(gstrConfigName, "OverrideLocation")  'Doesnt exist?  10/14/2013???
            'gstrConfigName = ""
            'If strOverrideLocation = "1" Then  'Doesnt exist?  10/14/2013???
                
                'strDefaultLocation = strLocation
            '    strPartLocation = objModel.GetCustomInfoValue(gstrConfigName, "PartLocation")

            '    If strPartLocation = "1" Then
                
            '        strLocation = "F" 'Northern

            '    ElseIf strPartLocation = "2" = True Then

            '        strLocation = "N" 'Nu
                
            '    ElseIf strPartLocation = "3" = True Then

            '        strLocation = "D" 'N2

            '    End If
            'Else
                'strLocation = strDefaultLocation
            'End If
            '*******************************************
            '*** ADDED 10/14/2013 TB Changes to lbKITPUR
            '*******************************************
            'strLocation = objModel.GetCustomInfoValue(gstrConfigName, "lbKITPUR")
            'force KITPUR to NKIT
            strLocation = "NKIT"
            'OP 10
            If strLocation = "KIT" Or strLocation = "NKIT" Or strLocation = "N2KIT" Then
                Print #1, QuoteMe(strPartNumber) & QuoteMe(strLocation) & "10 " & "0" & " " & "0" & " " & QuoteMe("COMMON SET") & "0"
            ElseIf strLocation = "PUR" Or strLocation = "NPUR" Or strLocation = "N2PUR" Then
                Print #1, QuoteMe(strPartNumber) & QuoteMe(strLocation) & "10 " & "0" & " " & "0" & " " & QuoteMe("COMMON SET") & "0"
            End If
            
            'OP 10
            'If strLocation = "F" Then
            '    Print #1, QuoteMe(strPartNumber) & QuoteMe("KIT") & "10 " & "0" & " " & "0" & " " & QuoteMe("COMMON SET") & "0"
            'ElseIf strLocation = "N" Then
            '    Print #1, QuoteMe(strPartNumber) & QuoteMe("NKIT") & "10 " & "0" & " " & "0" & " " & QuoteMe("COMMON SET") & "0"
            'End If
            
            'OP 20
            If objModel.GetCustomInfoValue(gstrConfigName, "OP20") <> "" Then
                strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "OP20")
                strSetup = objModel.GetCustomInfoValue(gstrConfigName, "OP20_S")
                strRun = objModel.GetCustomInfoValue(gstrConfigName, "OP20_R")
                If strSetup = "" Or strSetup = "." Then
                    strSetup = 0
                End If
                If strRun = "" Or strRun = "." Then
                    strRun = 0
                End If
                If Not IsNumeric(strSetup) Or Not IsNumeric(strRun) Then
                    MsgBox "Text value invalid...Check log"
                    boolLogError = True
                    Print #2, "Text value where a Numeric value is required :: " & strPartNumber & " :: OP20"
                End If
                
                Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & "20 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
            End If
            
            'OP 30
            If objModel.GetCustomInfoValue(gstrConfigName, "OP30") <> "" Then
                strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "OP30")
                strSetup = objModel.GetCustomInfoValue(gstrConfigName, "OP30_S")
                strRun = objModel.GetCustomInfoValue(gstrConfigName, "OP30_R")
                If strSetup = "" Or strSetup = "." Then
                    strSetup = 0
                End If
                If strRun = "" Or strRun = "." Then
                    strRun = 0
                End If
                If Not IsNumeric(strSetup) Or Not IsNumeric(strRun) Then
                    MsgBox "Text value invalid...Check log"
                    boolLogError = True
                    Print #2, "Text value where a Numeric value is required :: " & strPartNumber & " :: OP30"
                End If
                
                Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & "30 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
            End If
            
            'OP 40
            If objModel.GetCustomInfoValue(gstrConfigName, "OP40") <> "" Then
                strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "OP40")
                strSetup = objModel.GetCustomInfoValue(gstrConfigName, "OP40_S")
                strRun = objModel.GetCustomInfoValue(gstrConfigName, "OP40_R")
                If strSetup = "" Or strSetup = "." Then
                    strSetup = 0
                End If
                If strRun = "" Or strRun = "." Then
                    strRun = 0
                End If
                If Not IsNumeric(strSetup) Or Not IsNumeric(strRun) Then
                    MsgBox "Text value invalid...Check log"
                    boolLogError = True
                    Print #2, "Text value where a Numeric value is required :: " & strPartNumber & " :: OP40"
                End If
                    
                Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & "40 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
            End If
            
            'OP 50
            If objModel.GetCustomInfoValue(gstrConfigName, "OP50") <> "" Then
                strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "OP50")
                strSetup = objModel.GetCustomInfoValue(gstrConfigName, "OP50_S")
                strRun = objModel.GetCustomInfoValue(gstrConfigName, "OP50_R")
                If strSetup = "" Or strSetup = "." Then
                    strSetup = 0
                End If
                If strRun = "" Or strRun = "." Then
                    strRun = 0
                End If
                If Not IsNumeric(strSetup) Or Not IsNumeric(strRun) Then
                    MsgBox "Text value invalid...Check log"
                    boolLogError = True
                    Print #2, "Text value where a Numeric value is required :: " & strPartNumber & " :: OP50"
                End If
                
                Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & "50 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
            End If
            
            'OP 60
            If objModel.GetCustomInfoValue(gstrConfigName, "OP60") <> "" Then
                strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "OP60")
                strSetup = objModel.GetCustomInfoValue(gstrConfigName, "OP60_S")
                strRun = objModel.GetCustomInfoValue(gstrConfigName, "OP60_R")
                If strSetup = "" Or strSetup = "." Then
                    strSetup = 0
                End If
                If strRun = "" Or strRun = "." Then
                    strRun = 0
                End If
                If Not IsNumeric(strSetup) Or Not IsNumeric(strRun) Then
                    MsgBox "Text value invalid...Check log"
                    boolLogError = True
                    Print #2, "Text value where a Numeric value is required :: " & strPartNumber & " :: OP60"
                End If
                
                Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & "60 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
            End If
            
            'OP 70
            If objModel.GetCustomInfoValue(gstrConfigName, "OP70") <> "" Then
                strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "OP70")
                strSetup = objModel.GetCustomInfoValue(gstrConfigName, "OP70_S")
                strRun = objModel.GetCustomInfoValue(gstrConfigName, "OP70_R")
                If strSetup = "" Or strSetup = "." Then
                    strSetup = 0
                End If
                If strRun = "" Or strRun = "." Then
                    strRun = 0
                End If
                If Not IsNumeric(strSetup) Or Not IsNumeric(strRun) Then
                    MsgBox "Text value invalid...Check log"
                    boolLogError = True
                    Print #2, "Text value where a Numeric value is required :: " & strPartNumber & " :: OP70"
                End If
                
                Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & "70 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
            End If

            'OP 80
            If objModel.GetCustomInfoValue(gstrConfigName, "OP80") <> "" Then
                strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "OP80")
                strSetup = objModel.GetCustomInfoValue(gstrConfigName, "OP80_S")
                strRun = objModel.GetCustomInfoValue(gstrConfigName, "OP80_R")
                If strSetup = "" Or strSetup = "." Then
                    strSetup = 0
                End If
                If strRun = "" Or strRun = "." Then
                    strRun = 0
                End If
                If Not IsNumeric(strSetup) Or Not IsNumeric(strRun) Then
                    MsgBox "Text value invalid...Check log"
                    boolLogError = True
                    Print #2, "Text value where a Numeric value is required :: " & strPartNumber & " :: OP80"
                End If
                
                Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & "80 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
            End If
            
            'OP 90
            If objModel.GetCustomInfoValue(gstrConfigName, "OP90") <> "" Then
                strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "OP90")
                strSetup = objModel.GetCustomInfoValue(gstrConfigName, "OP90_S")
                strRun = objModel.GetCustomInfoValue(gstrConfigName, "OP90_R")
                If strSetup = "" Or strSetup = "." Then
                    strSetup = 0
                End If
                If strRun = "" Or strRun = "." Then
                    strRun = 0
                End If
                If Not IsNumeric(strSetup) Or Not IsNumeric(strRun) Then
                    MsgBox "Text value invalid...Check log"
                    boolLogError = True
                    Print #2, "Text value where a Numeric value is required :: " & strPartNumber & " :: OP90"
                End If
                
                Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & "90 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
            End If
            
            'OP 100
            If objModel.GetCustomInfoValue(gstrConfigName, "OP100") <> "" Then
                strWorkCenter = objModel.GetCustomInfoValue(gstrConfigName, "OP100")
                strSetup = objModel.GetCustomInfoValue(gstrConfigName, "OP100_S")
                strRun = objModel.GetCustomInfoValue(gstrConfigName, "OP100_R")
                If strSetup = "" Or strSetup = "." Then
                    strSetup = 0
                End If
                If strRun = "" Or strRun = "." Then
                    strRun = 0
                End If
                If Not IsNumeric(strSetup) Or Not IsNumeric(strRun) Then
                    MsgBox "Text value invalid...Check log"
                    boolLogError = True
                    Print #2, "Text value where a Numeric value is required :: " & strPartNumber & " :: OP100"
                End If
                
                Print #1, QuoteMe(strPartNumber) & QuoteMe(strWorkCenter) & "100 " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
            End If
            
            '******* ADDED 10/14/13 TB **************************
            '*** check for RevisitBeforeExportASM flag is set ***
            '****************************************************
            tbRevisitNote = ""
            bRevisitCB = False
            tbRevisitNote = objModel.GetCustomInfoValue(gstrConfigName, "RevisitNoteASM")
            If objModel.GetCustomInfoValue(gstrConfigName, "RevisitBeforeExportASM") = "Checked" Then
                bRevisitCB = True
                Print #2, "REVISIT *** " & strPartNumber & ".  Message *** " & tbRevisitNote
                boolLogMessage = True
            End If
            '***************************************************
            
            'Parent Outsourced Operation   ***ADDED 8/20/2014  TB ****  (copied from child route)

            strReqOS = objModel.GetCustomInfoValue(gstrConfigName, "ReqOS_A")
            If strReqOS = "Checked" Then    'Outsourced
                If objModel.GetCustomInfoValue(gstrConfigName, "OS_WC_A") = "" Then  'No OS Number input
                    boolLogError = True
                    Print #2, "Outsourced Operation not included for assembly " & strPartNumber
                    strOSNumber_A = ""
                Else
                    strOS_WC_A = Left(objModel.GetCustomInfoValue(gstrConfigName, "OS_WC_A"), 3) 'Left 3 chars of String
                    strOSNumber_A = strOS_WC_A & "-" & strPartNumber
                End If
                'Bypassing all OSLocation checks
                'If objModel.GetCustomInfoValue(gstrConfigName, "OSLocation_A") = "" Then  'No OS Location input
                '    boolLogError = True
                '    Print #2, "OSLocation not included for assmebly " & strPartNumber & ".  Defaulting to N1"
                '    strOSLocation_A = "0"
                'Else
                '    strOSLocation_A = objModel.GetCustomInfoValue(gstrConfigName, "OSLocation_A")
                'End If
                strOSLocation_A = 1
                
                If objModel.GetCustomInfoValue(gstrConfigName, "OS_OP_A") = "" Or objModel.GetCustomInfoValue(gstrConfigName, "OS_OP_A") = "0" Then   'No Outsource Operation Number input
                    boolLogError = True
                    Print #2, "Outsource Operation Number not included for assembly " & strPartNumber
                    strOS_OP_A = ""
                Else
                    strOS_OP_A = objModel.GetCustomInfoValue(gstrConfigName, "OS_OP_A")
                End If
                strSetup = 0
                strRun = 0
                If strOSLocation_A = "0" Then
                    strOSLocation_A = ""
                ElseIf strOSLocation_A = "1" Then
                    strOSLocation_A = "NU"
                ElseIf strOSLocation_A = "2" Then
                    strOSLocation_A = "N2"
                End If
                    
                If (strSetup <> "" And strRun <> "") Then
                    Print #1, QuoteMe(strPartNumber) & QuoteMe(strOS_WC_A & strOSLocation_A) & strOS_OP_A & " " & strSetup & " " & strRun & " " & QuoteMe("COMMON SET") & "0"
                End If
            End If
            
            
Print #1, " "
End Sub
