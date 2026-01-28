
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
            Case "rbOtherAlum"
                OtherMaterial.rbOtherAlum.value = True
            Case "rbBendTable"
                SemiAutoPilot.rbBendTable.value = True
            Case "rbKFactor"
                SemiAutoPilot.rbKFactor.value = True
            Case "tbKFactor"
                SemiAutoPilot.tbKFactor.value = Mid(sLineOfText, InStrRev(sLineOfText, ",") + 1, Len(sLineOfText) - InStrRev(sLineOfText, ","))
            Case "cbCreateDXF"
                SemiAutoPilot.cbCreateDXF.value = True
            Case "cbCreateDrawing"
                SemiAutoPilot.cbCreateDrawing.value = True
            Case "obDim"
                SemiAutoPilot.obDim.value = True
            Case "cbVisible"
                SemiAutoPilot.cbVisible.value = True
            Case "cbGrain"
                CustomProps.cbGrain.value = True
            Case "cbCommon"
                CustomProps.cbCommon.value = True
            Case "tbCustomer"
                CustomProps.tbCustomer.value = Mid(sLineOfText, InStrRev(sLineOfText, ",") + 1, Len(sLineOfText) - InStrRev(sLineOfText, ","))
            Case "tbPrint"
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

Public Sub SaveCurrentModel()
    If Not swModel Is Nothing Then
        swModel.Save3 1, 0, 0
        swModel.ClearSelection2 True
    End If
End Sub

' Additional file operation procedures can be added here