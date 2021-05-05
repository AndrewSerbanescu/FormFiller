# FormFiller

Dim i As Integer, j As Integer

Private Sub chkDPSP_Click()
    MultiPage1.Pages(2).Visible = chkDPSP.Value

    With Sheet1
        i = .Range("DPSP").Row
        j = .Range("RPP").Row - 2
        If chkDPSP.Value = False Then
            .Rows(i & ":" & j).Hidden = True
        Else
            .Rows(i & ":" & j).Hidden = False
        End If
    End With
End Sub

Private Sub chkPayroll_Click()
    fraPayroll.Enabled = chkPayroll.Value

    With Sheet1
        i = .Range("HRPayroll").Row
        j = .Range("PlanSponsorStatement").Row - 2
        If chkPayroll.Value = False Then
            .Rows(i & ":" & j).Hidden = True
        Else
            .Rows(i & ":" & j).Hidden = False
        End If
    End With
End Sub

Private Sub chkPortfolio_Click()
If chkPortfolio.Value = True Then
[595:597].EntireRow.Hidden = False
Else: [595:597].EntireRow.Hidden = True
End If
End Sub

Private Sub chkRPP_Click()
    MultiPage1.Pages(3).Visible = chkRPP.Value

    With Sheet1
        i = .Range("RPP").Row
        j = .Range("TFSA").Row - 2
        If chkRPP.Value = False Then
            .Rows(i & ":" & j).Hidden = True
        Else
            .Rows(i & ":" & j).Hidden = False
        End If
    End With
End Sub

Private Sub chkRRSP_Click()
    MultiPage1.Pages(1).Visible = chkRRSP.Value

    With Sheet1
        i = .Range("RRSP").Row
        j = .Range("DPSP").Row - 2
        If chkRRSP.Value = False Then
            .Rows(i & ":" & j).Hidden = True
        Else
            .Rows(i & ":" & j).Hidden = False
        End If
    End With
End Sub

Private Sub chkSPP_Click()
    MultiPage1.Pages(5).Visible = chkSPP.Value

    With Sheet1
        i = .Range("SPP").Row
        j = .Range("HRServices").Row - 2
        If chkSPP.Value = False Then
            .Rows(i & ":" & j).Hidden = True
        Else
            .Rows(i & ":" & j).Hidden = False
        End If
    End With
End Sub

Private Sub chkTFSA_Click()
    MultiPage1.Pages(4).Visible = chkTFSA.Value

    With Sheet1
        i = .Range("TFSA").Row
        j = .Range("SPP").Row - 2
        If chkTFSA.Value = False Then
            .Rows(i & ":" & j).Hidden = True
        Else
            .Rows(i & ":" & j).Hidden = False
        End If
    End With
End Sub

Private Sub cmdAddAffiliate_Click()
    If optAffiliateNo.Value = True Then
        MsgBox "You select no affiliate information required... Please change option", vbExclamation, "Error"
    Else
        frmAffiliate.Show vbModal
    End If
End Sub

Private Sub cmdAddClass_Click()
    If optClassNo.Value = True Then
        MsgBox "You select no class information required... Please change option", vbExclamation, "Error"
    Else
        frmClass.Show vbModal
    End If
End Sub

Private Sub cmdAddDivision_Click()
    If optDivisionalNo.Value = True Then
        MsgBox "You select no divisional information required... Please change option", vbExclamation, "Error"
    Else
        frmDivision.Show vbModal
    End If
End Sub

Private Sub cmdAddSecondary_Click()
    If optSecNo.Value = True Then
        MsgBox "You select no secondary contact required... Please change option", vbExclamation, "Error"
    Else
        frmSecondary.Show vbModal
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDPSPClose_Click()
    Unload Me
End Sub

Private Sub cmdDPSPEnterCClass_Click()
    If optCClassNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Class"
    frmDPSPCClass.Show vbModal
End Sub

Private Sub cmdDPSPEnterCDivision_Click()
    If optCDivisionNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Division"
    frmDPSPCClass.Show vbModal
End Sub

Private Sub cmdDPSPEnterEClass_Click()
    If optDPSPEClassNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Class"
    frmDPSPEClass.Show vbModal
End Sub

Private Sub cmdDPSPEnterEDivision_Click()
    If optDPSPEDivisionNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Division"
    frmDPSPEClass.Show vbModal
End Sub

Private Sub cmdDPSPEnterWClass_Click()
    If optWClassNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Class"
    frmDPSPWClass.Show vbModal
End Sub

Private Sub cmdDPSPEnterWDivision_Click()
    If optWDivisionNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Division"
    frmDPSPWClass.Show vbModal
End Sub

Private Sub cmdDPSPResetCClass_Click()
    With Sheet1
        .Range("DPSPContributionClass").Value = "Class"
        i = .Range("DPSPContributionClass").Row + 5
        j = .Range("DPSPContributionDivision").Row
        .Range("C" & i - 4 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdDPSPResetCDivision_Click()
    With Sheet1
        .Range("DPSPContributionDivision").Value = "Division"
        i = .Range("DPSPContributionDivision").Row + 5
        j = .Range("DPSPWithdrawals").Row
        .Range("C" & i - 4 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdDPSPResetEClass_Click()
    With Sheet1
        .Range("DPSPEligibilityClass").Value = "Class"
        i = .Range("DPSPEligibilityClass").Row + 5
        j = .Range("DPSPEligibilityDivision").Row
        .Range("C" & i - 4 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdDPSPResetEDivision_Click()
    With Sheet1
        .Range("DPSPEligibilityDivision").Value = "Division"
        i = .Range("DPSPEligibilityDivision").Row + 5
        j = .Range("DPSPContribution").Row
        .Range("C" & i - 4 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdDPSPResetWClass_Click()
    With Sheet1
        .Range("DPSPWithdrawalClass").Value = "Class"
        i = .Range("DPSPWithdrawalClass").Row + 7
        j = .Range("DPSPWithdrawalDivision").Row
        .Range("C" & i - 6 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdDPSPResetWDivision_Click()
    With Sheet1
        .Range("DPSPWithdrawalDivision").Value = "Division"
        i = .Range("DPSPWithdrawalDivision").Row + 7
        j = .Range("DPSPWithdrawalDivision").End(xlDown).Row
        .Range("C" & i - 6 & ":C" & i).ClearContents
        If i < j Then
            .Rows(i & ":" & j).Delete
        End If
    End With
End Sub

Private Sub cmdDPSPSave_Click()
    Dim r As Integer
    
    With Sheet1
        r = .Range("DPSP").Row
        .Range("C" & r + 1).Value = txtEffectiveDate1.Text
        .Range("C" & r + 2).Value = txtPolicyNo1.Text
        .Range("C" & r + 3).Value = txtPlanPurpose1.Text
        .Range("C" & r + 4).Value = txtOriginalPlan1.Text
        .Range("C" & r + 5).Value = txtRPlanName1.Text
        .Range("C" & r + 6).Value = txtEmpClass1.Text
        .Range("C" & r + 7).Value = txtFund1.Text
        .Range("C" & r + 8).Value = txtFundCode11.Text
        .Range("C" & r + 9).Value = txtFundCode21.Text
        
        r = .Range("DPSPEligibility").Row
        .Range("C" & r + 1).Value = txtDPSPE1.Text
        .Range("C" & r + 2).Value = txtDPSPE2.Text
        .Range("C" & r + 3).Value = txtDPSPE3.Text
        .Range("C" & r + 4).Value = txtDPSPE4.Text
        .Range("C" & r + 5).Value = IIf(optDPSPEClassYes.Value = True, "Yes", "No")
        .Range("C" & r + 6).Value = IIf(optDPSPEDivisionYes.Value = True, "Yes", "No")
        
        r = .Range("DPSPContribution").Row
        .Range("C" & r + 1).Value = txtDPSPC1.Text
        .Range("C" & r + 2).Value = txtDPSPC2.Text
        .Range("C" & r + 3).Value = txtDPSPC3.Text
        .Range("C" & r + 4).Value = txtDPSPC4.Text
        .Range("C" & r + 5).Value = txtDPSPC5.Text
        .Range("C" & r + 6).Value = IIf(optDPSPCClassYes.Value = True, "Yes", "No")
        .Range("C" & r + 7).Value = IIf(optDPSPCDivisionYes.Value = True, "Yes", "No")
        .Range("C" & r + 8).Value = txtDPSPC8.Text
        .Range("C" & r + 9).Value = txtDPSPC10.Text
    
        r = .Range("DPSPWithdrawals").Row
        .Range("C" & r + 1).Value = txtDPSPW1.Text
        .Range("C" & r + 2).Value = txtDPSPW2.Text
        .Range("C" & r + 3).Value = txtDPSPW3.Text
        .Range("C" & r + 4).Value = IIf(optDPSPWClassYes.Value = True, "Yes", "N0")
        .Range("C" & r + 5).Value = IIf(optDPSPWDivisionYes.Value = True, "Yes", "No")
    
        r = .Range("DPSPTable").Row
        .Range("B" & r + 1).Value = txtDPSPV1.Text
        .Range("C" & r + 1).Value = txtDPSPV2.Text
        .Range("D" & r + 1).Value = txtDPSPV3.Text
        .Range("B" & r + 2).Value = txtDPSPV4.Text
        .Range("C" & r + 2).Value = txtDPSPV5.Text
        .Range("D" & r + 2).Value = txtDPSPV6.Text
        .Range("B" & r + 3).Value = txtDPSPV7.Text
        .Range("C" & r + 3).Value = txtDPSPV8.Text
        .Range("D" & r + 3).Value = txtDPSPV9.Text
    End With
End Sub

Private Sub cmdEnterCClass_Click()
    If optCClassNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Class"
    frmRRSPCClass.Show vbModal
End Sub

Private Sub cmdEnterCDivision_Click()
    If optCDivisionNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Division"
    frmRRSPCClass.Show vbModal
End Sub

Private Sub cmdEnterEClass_Click()
    If optEClassNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Class"
    frmRRSPEClass.Show vbModal
End Sub

Private Sub cmdEnterEDivision_Click()
    If optEDivisionNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Division"
    frmRRSPEClass.Show vbModal
End Sub

Private Sub cmdEnterPayrollClass_Click()
    txtEligibilityType.Text = "Class"
    frmPayrollClass.Show vbModal
End Sub

Private Sub cmdEnterPayrollDivision_Click()
    txtEligibilityType.Text = "Division"
    frmPayrollClass.Show vbModal
End Sub

Private Sub cmdEnterWClass_Click()
    If optWClassNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Class"
    frmRRSPWClass.Show vbModal
End Sub

Private Sub cmdEnterWDivision_Click()
    If optWDivisionNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Division"
    frmRRSPWClass.Show vbModal
End Sub

Private Sub cmdResetAffiliate_Click()
    Dim c As Range
    Dim cRow As Integer
    
    With Sheet1
        Set c = .UsedRange.Find(What:="Affiliated company Information", LookIn:=xlValues)
        If Not c Is Nothing Then
            i = .Range("Z" & c.Row).Value
            Select Case i
            Case 1:
                .Range("C" & c.Row + 1 & ":C" & c.Row + 5).ClearContents
                .Range("Z" & c.Row).Value = 1
            Case Is > 1:
                cRow = c.Row
                Set c = .UsedRange.Find(What:="Plan Administrator information and level of access:", LookIn:=xlValues)
                .Rows(cRow + 6 & ":" & c.Row - 1).Delete
                .Range("C" & cRow + 1 & ":C" & cRow + 5).ClearContents
                .Range("Z" & cRow).Value = 1
            End Select
            
            MsgBox "Affiliate section has been reset", vbInformation, "Success"
        Else
            MsgBox "Affiliate section is not present", vbExclamation, "Error"
        End If
    End With
End Sub

Private Sub cmdResetClass_Click()
    Dim c As Range
    
    With Sheet1
        Set c = .UsedRange.Find(What:="Classes (if applicable)", LookIn:=xlValues)
        If Not c Is Nothing Then
            i = .Range("B" & c.Row + 2).End(xlToRight).Column
            Select Case i
            Case 3:
                .Range("C" & c.Row + 2).ClearContents
            Case Is > 3:
                .Range(.Cells(c.Row + 2, 3), .Cells(c.Row + 2, i)).ClearContents
            End Select
            
            cboEClass.Clear
            MsgBox "Class section has been reset", vbInformation, "Success"
        Else
            MsgBox "Class section is not present", vbExclamation, "Error"
        End If
    End With
End Sub

Private Sub cmdResetDivision_Click()
    Dim c As Range
    Dim cRow As Integer
    
    With Sheet1
        Set c = .UsedRange.Find(What:="Divisional (if applicable)", LookIn:=xlValues)
        If Not c Is Nothing Then
            i = .Range("Z" & c.Row).Value
            Select Case i
            Case 1:
                .Range("C" & c.Row + 1 & ":C" & c.Row + 5).Value = "N/A"
                .Range("Z" & c.Row).Value = 1
            Case Is > 1:
                .Range("C" & c.Row + 1 & ":C" & c.Row + 5).Value = "N/A"
                .Range("Z" & c.Row).Value = 1
                cRow = c.Row
                Set c = .UsedRange.Find(What:="Classes (if applicable)", LookIn:=xlValues)
                .Rows(cRow + 6 & ":" & c.Row - 1).Delete
            End Select
            
            cboEDivision.Clear
            MsgBox "Divisional section has been reset", vbInformation, "Success"
        Else
            MsgBox "Divisional section is not present", vbExclamation, "Error"
        End If
    End With
End Sub

Private Sub cmdResetSecondary_Click()
    Dim c As Range
    Dim cRow As Integer
    
    With Sheet1
        Set c = .UsedRange.Find(What:="Secondary Contact (if applicable)", LookIn:=xlValues)
        If Not c Is Nothing Then
            i = .Range("Z" & c.Row).Value
            Select Case i
            Case 1:
                .Range("C" & c.Row + 1 & ":C" & c.Row + 3).ClearContents
                .Range("Z" & c.Row).Value = 1
            Case Is > 1:
                .Range("C" & c.Row + 1 & ":C" & c.Row + 3).ClearContents
                .Range("Z" & c.Row).Value = 1
                cRow = c.Row
                Set c = .UsedRange.Find(What:="Divisional (if applicable)", LookIn:=xlValues)
                .Rows(cRow + 4 & ":" & c.Row - 1).Delete
            End Select
            
            MsgBox "Secondary contact section has been reset", vbInformation, "Success"
        Else
            MsgBox "Secondary contact section is not present", vbExclamation, "Error"
        End If
    End With
End Sub

Private Sub cmdRPPEnterCClass_Click()
    If optRPPCClassNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Class"
    frmRPPCClass.Show vbModal
End Sub

Private Sub cmdRPPEnterCDivision_Click()
    If optRPPEDivisionNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Division"
    frmRPPCClass.Show vbModal
End Sub

Private Sub cmdRPPEnterEClass_Click()
    If optRPPEClassNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Class"
    frmRPPEClass.Show vbModal
End Sub

Private Sub cmdRPPEnterEDivision_Click()
    If optRPPEDivisionNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Division"
    frmRPPEClass.Show vbModal
End Sub

Private Sub cmdRPPEnterWClass_Click()
    If optRPPWClassNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Class"
    frmRPPWClass.Show vbModal
End Sub

Private Sub cmdRPPEnterWDivision_Click()
    If optRPPWDivisionNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Division"
    frmRPPWClass.Show vbModal
End Sub

Private Sub cmdRPPResetCClass_Click()
    With Sheet1
        .Range("RPPContributionClass").Value = "Class"
        i = .Range("RPPContributionClass").Row + 5
        j = .Range("RPPContributionDivision").Row
        .Range("C" & i - 4 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdRPPResetCDivision_Click()
    With Sheet1
        .Range("RPPContributionDivision").Value = "Division"
        i = .Range("RPPContributionDivision").Row + 5
        j = .Range("RPPWithdrawals").Row
        .Range("C" & i - 4 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdRPPResetEClass_Click()
    With Sheet1
        .Range("RPPEligibilityClass").Value = "Class"
        i = .Range("RPPEligibilityClass").Row + 5
        j = .Range("RPPEligibilityDivision").Row
        .Range("C" & i - 4 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdRPPResetEDivision_Click()
    With Sheet1
        .Range("RPPEligibilityDivision").Value = "Division"
        i = .Range("RPPEligibilityDivision").Row + 5
        j = .Range("RPPContribution").Row
        .Range("C" & i - 4 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdRPPResetWClass_Click()
    With Sheet1
        .Range("RPPWithdrawalClass").Value = "Class"
        i = .Range("RPPWithdrawalClass").Row + 4
        j = .Range("RPPWithdrawalDivision").Row
        .Range("C" & i - 3 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdRPPResetWDivision_Click()
    With Sheet1
        .Range("RPPWithdrawalDivision").Value = "Division"
        i = .Range("RPPWithdrawalDivision").Row + 4
        j = .Range("RPPWithdrawalDivision").End(xlDown).Row
        .Range("C" & i - 3 & ":C" & i).ClearContents
        If i < j Then
            .Rows(i & ":" & j).Delete
        End If
    End With
End Sub

Private Sub cmdRPPSave_Click()
    Dim r As Integer
    
    With Sheet1
        r = .Range("RPP").Row
        .Range("C" & r + 1).Value = txtRPP1.Text
        .Range("C" & r + 2).Value = txtRPP2.Text
        .Range("C" & r + 3).Value = txtRPP3.Text
        .Range("C" & r + 4).Value = txtRPP4.Text
        .Range("C" & r + 5).Value = txtRPP5.Text
        .Range("C" & r + 6).Value = txtRPP6.Text
        .Range("C" & r + 7).Value = txtRPP7.Text
        .Range("C" & r + 8).Value = txtRPP8.Text
        .Range("C" & r + 9).Value = txtRPP9.Text
        
        r = .Range("RPPEligibility").Row
        .Range("C" & r + 1).Value = txtRPPE1.Text
        .Range("C" & r + 2).Value = txtRPPE2.Text
        .Range("C" & r + 3).Value = txtRPPE3.Text
        .Range("C" & r + 4).Value = txtRPPE4.Text
        .Range("C" & r + 5).Value = IIf(optRPPEClassYes.Value = True, "Yes", "No")
        .Range("C" & r + 6).Value = IIf(optRPPEDivisionYes.Value = True, "Yes", "No")
        
        r = .Range("RPPContribution").Row
        .Range("C" & r + 1).Value = txtRPPC1.Text
        .Range("C" & r + 2).Value = txtRPPC2.Text
        .Range("C" & r + 3).Value = txtRPPC3.Text
        .Range("C" & r + 4).Value = txtRPPC4.Text
        .Range("C" & r + 5).Value = txtRPPC5.Text
        .Range("C" & r + 6).Value = IIf(optRPPCClassYes.Value = True, "Yes", "No")
        .Range("C" & r + 7).Value = IIf(optRPPCDivisionYes.Value = True, "Yes", "No")
        .Range("C" & r + 8).Value = txtRPPC7.Text
        .Range("C" & r + 9).Value = txtRPPC8.Text
        .Range("C" & r + 10).Value = txtRPPC9.Text
        .Range("C" & r + 11).Value = txtRPPC10.Text
    
        r = .Range("RPPWithdrawals").Row
        .Range("C" & r + 1).Value = txtRPPW1.Text
        .Range("C" & r + 2).Value = txtRPPW2.Text
        .Range("C" & r + 3).Value = txtRPPW3.Text
        .Range("C" & r + 4).Value = IIf(optRPPWClassYes.Value = True, "Yes", "N0")
        .Range("C" & r + 5).Value = IIf(optRPPWDivisionYes.Value = True, "Yes", "No")
    
        r = .Range("RPPTable").Row
        .Range("B" & r + 1).Value = txtRPPV1.Text
        .Range("C" & r + 1).Value = txtRPPV2.Text
        .Range("D" & r + 1).Value = txtRPPV3.Text
        .Range("B" & r + 2).Value = txtRPPV4.Text
        .Range("C" & r + 2).Value = txtRPPV5.Text
        .Range("D" & r + 2).Value = txtRPPV6.Text
        .Range("B" & r + 3).Value = txtRPPV7.Text
        .Range("C" & r + 3).Value = txtRPPV8.Text
        .Range("D" & r + 3).Value = txtRPPV9.Text
    End With
End Sub

Private Sub cmdRRSPClose_Click()
    Unload Me
End Sub

Private Sub cmdRRSPResetCClass_Click()
    With Sheet1
        .Range("RRSPContributionClass").Value = "Class"
        i = .Range("RRSPContributionClass").Row + 5
        j = .Range("RRSPContributionDivision").Row
        .Range("C" & i - 4 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdRRSPResetCDivision_Click()
    With Sheet1
        .Range("RRSPContributionDivision").Value = "Division"
        i = .Range("RRSPContributionDivision").Row + 5
        j = .Range("RRSPWithdrawals").Row
        .Range("C" & i - 4 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdRRSPResetEClass_Click()
    With Sheet1
        .Range("RRSPEligibilityClass").Value = "Class"
        i = .Range("RRSPEligibilityClass").Row + 5
        j = .Range("RRSPEligibilityDivision").Row
        .Range("C" & i - 4 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdRRSPResetEDivision_Click()
    With Sheet1
        .Range("RRSPEligibilityDivision").Value = "Division"
        i = .Range("RRSPEligibilityDivision").Row + 5
        j = .Range("RRSPContribution").Row
        .Range("C" & i - 4 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdRRSPResetWClass_Click()
    With Sheet1
        .Range("RRSPWithdrawalClass").Value = "Class"
        i = .Range("RRSPWithdrawalClass").Row + 7
        j = .Range("RRSPWithdrawalDivision").Row
        .Range("C" & i - 6 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdRRSPResetWDivision_Click()
    With Sheet1
        .Range("RRSPWithdrawalDivision").Value = "Division"
        i = .Range("RRSPWithdrawalDivision").Row + 7
        j = .Range("RRSPWithdrawalDivision").End(xlDown).Row
        .Range("C" & i - 6 & ":C" & i).ClearContents
        If i < j Then
            .Rows(i & ":" & j).Delete
        End If
    End With
End Sub

Private Sub cmdRRSPSave_Click()
    Dim r As Integer
    
    With Sheet1
        r = .Range("RRSP").Row
        .Range("C" & r + 1).Value = txtEffectiveDate.Text
        .Range("C" & r + 2).Value = txtPolicyNo.Text
        .Range("C" & r + 3).Value = txtPlanPurpose.Text
        .Range("C" & r + 4).Value = txtOriginalPlan.Text
        .Range("C" & r + 5).Value = txtRPlanName.Text
        .Range("C" & r + 6).Value = txtEmpClass.Text
        .Range("C" & r + 7).Value = txtFund.Text
        .Range("C" & r + 8).Value = txtFundCode1.Text
        .Range("C" & r + 9).Value = txtFundCode2.Text
        .Range("C" & r + 10).Value = txtDefinition.Text
        
        r = .Range("RRSPEligibility").Row
        .Range("C" & r + 1).Value = txtE1.Text
        .Range("C" & r + 2).Value = txtE2.Text
        .Range("C" & r + 3).Value = txtE3.Text
        .Range("C" & r + 4).Value = txtE4.Text
        .Range("C" & r + 5).Value = IIf(optEClassYes.Value = True, "Yes", "No")
        .Range("C" & r + 6).Value = IIf(optEDivisionYes.Value = True, "Yes", "No")
        
        r = .Range("RRSPContribution").Row
        .Range("C" & r + 1).Value = txtC1.Text
        .Range("C" & r + 2).Value = txtC2.Text
        .Range("C" & r + 3).Value = txtC3.Text
        .Range("C" & r + 4).Value = txtC4.Text
        .Range("C" & r + 5).Value = txtC5.Text
        .Range("C" & r + 6).Value = IIf(optCClassYes.Value = True, "Yes", "No")
        .Range("C" & r + 7).Value = IIf(optCDivisionYes.Value = True, "Yes", "No")
        .Range("C" & r + 8).Value = txtC6.Text
        .Range("C" & r + 9).Value = txtC7.Text
        .Range("C" & r + 10).Value = txtC8.Text
    
        r = .Range("RRSPWithdrawals").Row
        .Range("C" & r + 1).Value = txtW1.Text
        .Range("C" & r + 2).Value = txtW2.Text
        .Range("C" & r + 3).Value = txtW3.Text
        .Range("C" & r + 4).Value = IIf(optWClassYes.Value = True, "Yes", "N0")
        .Range("C" & r + 5).Value = IIf(optWDivisionYes.Value = True, "Yes", "No")
    End With
End Sub

Private Sub cmdRSSEnterEClass_Click()
    If optSPPEClassNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Class"
    frmSPPEClass.Show vbModal
End Sub

Private Sub cmdRSSEnterEDivision_Click()
    If optSPPEDivisionNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Division"
    frmSPPEClass.Show vbModal
End Sub

Private Sub cmdRSSResetEClass_Click()
    With Sheet1
        .Range("SPPEligibilityClass").Value = "Class"
        i = .Range("SPPEligibilityClass").Row + 5
        j = .Range("SPPEligibilityDivision").Row
        .Range("C" & i - 4 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdRSSResetEDivision_Click()
    With Sheet1
        .Range("SPPEligibilityDivision").Value = "Division"
        i = .Range("SPPEligibilityDivision").Row + 5
        j = .Range("SPPContribution").Row
        .Range("C" & i - 4 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdSaveMain_Click()
    Dim c As Range, d As Range
    Dim rowStart As Integer
    
    With Sheet1
        .Range("C3").Value = txtCompName.Value
        .Range("C4").Value = txtCompAdd.Value
        .Range("C5").Value = txtCompNum.Value
        .Range("C6").Value = cmbOrgType.Text
        .Range("C7").Value = txtBizNum.Value
        .Range("C9").Value = txtIncorp.Value
        .Range("C10").Value = txtFiscEnd.Value
        
        Set c = .UsedRange.Find(What:="Plan Administrator information and level of access: ", LookIn:=xlValues, LookAt:=xlPart)
        If Not c Is Nothing Then
            .Range("C" & c.Row + 1).Value = txtAdminName.Value
            .Range("C" & c.Row + 2).Value = txtAdminAdd.Value
            .Range("C" & c.Row + 3).Value = txtAdminEmail.Value
        End If
        
        '-------------------------------------------------------------------------------------------------------------
        'hide affiliate company section if not required
        Set c = .UsedRange.Find(What:="Company fiscal year end:", LookIn:=xlValues)
        rowStart = c.Row + 1
        
        If optAffiliateNo.Value = True Then
            Set c = .UsedRange.Find(What:="Plan Administrator information and level of access:", LookIn:=xlValues)
            .Rows(rowStart & ":" & c.Row - 1).Hidden = True
        Else
            Set c = .UsedRange.Find(What:="Plan Administrator information and level of access:", LookIn:=xlValues)
            .Rows(rowStart & ":" & c.Row).Hidden = False
        End If
        
        '-------------------------------------------------------------------------------------------------------------
        'hide divisional information section if not required
        Set c = .UsedRange.Find(What:="Divisional (if applicable)", LookIn:=xlValues)
        rowStart = c.Row + 6
        
        If optDivisionalNo.Value = True Then
            .Range("C" & c.Row + 5).Value = "No"
            '.Range("C" & c.Row + 1 & ":C" & c.Row + 4).Value = "N/A"
            Set c = .UsedRange.Find(What:="Classes (if applicable)", LookIn:=xlValues)
            If rowStart < c.Row - 1 Then .Rows(rowStart & ":" & c.Row - 1).Hidden = True
        Else
            .Range("C" & c.Row + 5).Value = "Yes"
            Set c = .UsedRange.Find(What:="Classes (if applicable)", LookIn:=xlValues)
            If rowStart < c.Row - 1 Then .Rows(rowStart & ":" & c.Row - 1).Hidden = False
        End If
        
        '-------------------------------------------------------------------------------------------------------------
        'hide class information section if not required
        Set c = .UsedRange.Find(What:="Classes (if applicable)", LookIn:=xlValues)
        rowStart = c.Row + 2
        
        If optClassNo.Value = True Then
            .Range("C" & c.Row + 1).Value = "No"
            .Rows(rowStart).Hidden = True
        Else
            .Range("C" & c.Row + 1).Value = "Yes"
            .Rows(rowStart).Hidden = False
        End If
        
        '-------------------------------------------------------------------------------------------------------------
        'hide secondary section if not required
        Set c = .UsedRange.Find(What:="Is a Secondary Contact required?", LookIn:=xlValues)
        rowStart = c.Row + 1
        
        If optSecNo.Value = True Then
            .Range("C" & c.Row).Value = "No"
            Set c = .UsedRange.Find(What:="Divisional (if applicable)", LookIn:=xlValues)
            .Rows(rowStart & ":" & c.Row - 1).Hidden = True
        Else
            .Range("C" & c.Row).Value = "Yes"
            Set c = .UsedRange.Find(What:="Divisional (if applicable)", LookIn:=xlValues)
            .Rows(rowStart & ":" & c.Row).Hidden = False
        End If
        
        '-------------------------------------------------------------------------------------------------------------
        'Save advisor information
        Set c = .UsedRange.Find(What:="Advisor Information and Advisor Access (if applicable)", LookIn:=xlValues)
        .Range("C" & c.Row + 1).Value = txtAdvisorName.Text
        .Range("C" & c.Row + 2).Value = txtAdvisorContact.Text
        .Range("C" & c.Row + 3).Value = txtAdvisorCode.Text
        .Range("C" & c.Row + 4).Value = txtAdvisorPlan.Text
        .Range("C" & c.Row + 5).Value = txtAdvisorInfo.Text
        .Range("C" & c.Row + 6).Value = txtAdvisorMember.Text
        .Range("C" & c.Row + 7).Value = txtAdvisorBenefit.Text
        
        '-------------------------------------------------------------------------------------------------------------
        'Save transfer details information
        Set c = .UsedRange.Find(What:="Transfer Details (If Applicable)", LookIn:=xlValues)
        .Range("C" & c.Row + 1).Value = txtT1.Text
        .Range("C" & c.Row + 2).Value = txtT2.Text
        .Range("C" & c.Row + 3).Value = txtT3.Text
        .Range("C" & c.Row + 4).Value = txtT4.Text
        .Range("C" & c.Row + 5).Value = txtT5.Text
        .Range("C" & c.Row + 6).Value = txtT6.Text
        .Range("C" & c.Row + 7).Value = txtT7.Text
        .Range("C" & c.Row + 8).Value = txtT8.Text
        .Range("C" & c.Row + 9).Value = txtT9.Text
        .Range("C" & c.Row + 10).Value = txtT10.Text
        .Range("C" & c.Row + 11).Value = txtT11.Text
        .Range("C" & c.Row + 12).Value = txtT12.Text
        
        '-------------------------------------------------------------------------------------------------------------
        'Save HR information
        Set c = .UsedRange.Find(What:="HR Services", LookIn:=xlValues)
        .Range("C" & c.Row + 1).Value = txtHR1.Text
        .Range("C" & c.Row + 2).Value = txtHR2.Text
        .Range("C" & c.Row + 3).Value = IIf(optBonusYes.Value = True, "Yes", "No")
        .Range("C" & c.Row + 4).Value = IIf(optPlanRightYes.Value = True, "Yes", "No")
        .Range("C" & c.Row + 5).Value = IIf(optHRYes.Value = True, "Yes", "No")
        
        .Range("C" & c.Row + 6).Value = txtHR15.Text
        
        Set c = .Range("HRPayroll")
        .Range("C" & c.Row + 2).Value = txtPayroll1.Text
        .Range("C" & c.Row + 3).Value = txtPayroll2.Text
        .Range("C" & c.Row + 4).Value = txtPayroll3.Text
        .Range("C" & c.Row + 5).Value = txtPayroll4.Text
        .Range("C" & c.Row + 6).Value = txtPayroll5.Text
        .Range("C" & c.Row + 8).Value = txtPayroll6.Text
        .Range("C" & c.Row + 9).Value = txtPayroll7.Text
        .Range("C" & c.Row + 10).Value = txtPayroll8.Text
        .Range("C" & c.Row + 11).Value = txtPayroll9.Text
        .Range("C" & c.Row + 12).Value = txtPayroll10.Text
    
        .Range("C1").EntireColumn.AutoFit
    End With
End Sub

Private Sub cmdSPPClose_Click()
    Unload Me
End Sub

Private Sub cmdSPPEnterCClass_Click()
    If optSPPCClassNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Class"
    frmSPPCClass.Show vbModal
End Sub

Private Sub cmdSPPEnterCDivision_Click()
    If optSPPCDivisionNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Division"
    frmSPPCClass.Show vbModal
End Sub

Private Sub cmdSPPEnterWClass_Click()
    If optSPPWClassNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Class"
    frmSPPWClass.Show vbModal
End Sub

Private Sub cmdSPPEnterWDivision_Click()
    If optSPPWDivisionNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Division"
    frmSPPWClass.Show vbModal
End Sub

Private Sub cmdSPPResetCClass_Click()
    With Sheet1
        .Range("SPPContributionClass").Value = "Class"
        i = .Range("SPPContributionClass").Row + 5
        j = .Range("SPPContributionDivision").Row
        .Range("C" & i - 4 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdSPPResetCDivision_Click()
    With Sheet1
        .Range("SPPContributionDivision").Value = "Division"
        i = .Range("SPPContributionDivision").Row + 5
        j = .Range("SPPWithdrawals").Row
        .Range("C" & i - 4 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdSPPResetWClass_Click()
    With Sheet1
        .Range("SPPWithdrawalClass").Value = "Class"
        i = .Range("SPPWithdrawalClass").Row + 7
        j = .Range("SPPWithdrawalDivision").Row
        .Range("C" & i - 6 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdSPPResetWDivision_Click()
    With Sheet1
        If .Rows(.Range("SPPWithdrawalDivision").Row).Hidden = True Then Exit Sub
        .Range("SPPWithdrawalDivision").Value = "Division"
        i = .Range("SPPWithdrawalDivision").Row + 7
        j = .Range("SPPWithdrawalDivision").End(xlDown).Row
        .Range("C" & i - 6 & ":C" & i).ClearContents
        If i < j Then
            .Rows(i & ":" & j).Delete
        End If
    End With
End Sub

Private Sub cmdSPPSave_Click()
    Dim r As Integer
    
    With Sheet1
        r = .Range("SPP").Row
        .Range("C" & r + 1).Value = txtSPP1.Text
        .Range("C" & r + 2).Value = txtSPP2.Text
        .Range("C" & r + 3).Value = txtSPP3.Text
        .Range("C" & r + 4).Value = txtSPP4.Text
        .Range("C" & r + 5).Value = txtSPP5.Text
        .Range("C" & r + 6).Value = txtSPP6.Text
        .Range("C" & r + 7).Value = txtSPP7.Text
        .Range("C" & r + 8).Value = txtSPP8.Text
        .Range("C" & r + 9).Value = txtSPP9.Text
        
        r = .Range("SPPEligibility").Row
        .Range("C" & r + 1).Value = txtSPPE1.Text
        .Range("C" & r + 2).Value = txtSPPE2.Text
        .Range("C" & r + 3).Value = txtSPPE3.Text
        .Range("C" & r + 4).Value = txtSPPE4.Text
        .Range("C" & r + 5).Value = IIf(optSPPEClassYes.Value = True, "Yes", "No")
        .Range("C" & r + 6).Value = IIf(optSPPEDivisionYes.Value = True, "Yes", "No")
        
        r = .Range("SPPContribution").Row
        .Range("C" & r + 1).Value = txtSPPC1.Text
        .Range("C" & r + 2).Value = txtSPPC2.Text
        .Range("C" & r + 3).Value = txtSPPC3.Text
        .Range("C" & r + 4).Value = txtSPPC4.Text
        .Range("C" & r + 5).Value = txtSPPC5.Text
        .Range("C" & r + 6).Value = IIf(optSPPCClassYes.Value = True, "Yes", "No")
        .Range("C" & r + 7).Value = IIf(optSPPCDivisionYes.Value = True, "Yes", "No")
        .Range("C" & r + 8).Value = txtSPPC6.Text
        .Range("C" & r + 9).Value = txtSPPC7.Text
        .Range("C" & r + 10).Value = txtSPPC8.Text
        .Range("C" & r + 11).Value = txtSPPC9.Text
        .Range("C" & r + 12).Value = txtSPPC10.Text
    
        r = .Range("SPPWithdrawals").Row
        .Range("C" & r + 1).Value = txtSPPW1.Text
        .Range("C" & r + 2).Value = txtSPPW2.Text
        .Range("C" & r + 3).Value = txtSPPW3.Text
        .Range("C" & r + 4).Value = IIf(optSPPWClassYes.Value = True, "Yes", "N0")
        .Range("C" & r + 5).Value = IIf(optSPPWDivisionYes.Value = True, "Yes", "No")
    
        r = .Range("SPPTable").Row
        .Range("B" & r + 1).Value = txtSPPV1.Text
        .Range("C" & r + 1).Value = txtSPPV2.Text
        .Range("D" & r + 1).Value = txtSPPV3.Text
        .Range("B" & r + 2).Value = txtSPPV4.Text
        .Range("C" & r + 2).Value = txtSPPV5.Text
        .Range("D" & r + 2).Value = txtSPPV6.Text
        .Range("B" & r + 3).Value = txtSPPV7.Text
        .Range("C" & r + 3).Value = txtSPPV8.Text
        .Range("D" & r + 3).Value = txtSPPV9.Text
    End With
End Sub

Private Sub cmdTFSAClose_Click()
    Unload Me
End Sub

Private Sub cmdTFSAEnterCClass_Click()
    If optTFSACClassNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Class"
    frmTFSACClass.Show vbModal
End Sub

Private Sub cmdTFSAEnterCDivision_Click()
    If optTFSACDivisionNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Division"
    frmTFSACClass.Show vbModal
End Sub

Private Sub cmdTFSAEnterEClass_Click()
    If optTFSAEClassNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Class"
    frmTFSAEClass.Show vbModal
End Sub

Private Sub cmdTFSAEnterEDivision_Click()
    If optTFSAEDivisionNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Division"
    frmTFSAEClass.Show vbModal
End Sub

Private Sub cmdTFSAEnterWClass_Click()
    If optTFSAWClassNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Class"
    frmTFSAWClass.Show vbModal
End Sub

Private Sub cmdTFSAEnterWDivision_Click()
    If optTFSAWDivisionNo.Value = True Then
        MsgBox "Provisions must be YES", vbExclamation, "Error"
        Exit Sub
    End If
    txtEligibilityType.Text = "Division"
    frmTFSAWClass.Show vbModal
End Sub

Private Sub cmdTFSAResetCClass_Click()
    With Sheet1
        .Range("TFSAContributionClass").Value = "Class"
        i = .Range("TFSAContributionClass").Row + 5
        j = .Range("TFSAContributionDivision").Row
        .Range("C" & i - 4 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdTFSAResetCDivision_Click()
    With Sheet1
        .Range("TFSAContributionDivision").Value = "Division"
        i = .Range("TFSAContributionDivision").Row + 5
        j = .Range("TFSAWithdrawals").Row
        .Range("C" & i - 4 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdTFSAResetEClass_Click()
    With Sheet1
        .Range("TFSAEligibilityClass").Value = "Class"
        i = .Range("TFSAEligibilityClass").Row + 5
        j = .Range("TFSAEligibilityDivision").Row
        .Range("C" & i - 4 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdTFSAResetEDivision_Click()
    With Sheet1
        .Range("TFSAEligibilityDivision").Value = "Division"
        i = .Range("TFSAEligibilityDivision").Row + 5
        j = .Range("TFSAContribution").Row
        .Range("C" & i - 4 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdTFSAResetWClass_Click()
    With Sheet1
        .Range("TFSAWithdrawalClass").Value = "Class"
        i = .Range("TFSAWithdrawalClass").Row + 7
        j = .Range("TFSAWithdrawalDivision").Row
        .Range("C" & i - 6 & ":C" & i - 1).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub cmdTFSAResetWDivision_Click()
    With Sheet1
        .Range("TFSAWithdrawalDivision").Value = "Division"
        i = .Range("TFSAWithdrawalDivision").Row + 7
        j = .Range("TFSAWithdrawalDivision").End(xlDown).Row
        .Range("C" & i - 6 & ":C" & i).ClearContents
        If i < j Then
            .Rows(i & ":" & j).Delete
        End If
    End With
End Sub

Private Sub cmdTFSASave_Click()
    Dim r As Integer
    
    With Sheet1
        r = .Range("TFSA").Row
        .Range("C" & r + 1).Value = txtTFSA1.Text
        .Range("C" & r + 2).Value = txtTFSA2.Text
        .Range("C" & r + 3).Value = txtTFSA3.Text
        .Range("C" & r + 4).Value = txtTFSA4.Text
        .Range("C" & r + 5).Value = txtTFSA5.Text
        .Range("C" & r + 6).Value = txtTFSA6.Text
        .Range("C" & r + 7).Value = txtTFSA7.Text
        .Range("C" & r + 8).Value = txtTFSA8.Text
        .Range("C" & r + 9).Value = txtTFSA9.Text
        
        r = .Range("TFSAEligibility").Row
        .Range("C" & r + 1).Value = txtTFSAE1.Text
        .Range("C" & r + 2).Value = txtTFSAE2.Text
        .Range("C" & r + 3).Value = txtTFSAE3.Text
        .Range("C" & r + 4).Value = txtTFSAE4.Text
        .Range("C" & r + 5).Value = IIf(optTFSAEClassYes.Value = True, "Yes", "No")
        .Range("C" & r + 6).Value = IIf(optTFSAEDivisionYes.Value = True, "Yes", "No")
        
        r = .Range("TFSAContribution").Row
        .Range("C" & r + 1).Value = txtTFSAC1.Text
        .Range("C" & r + 2).Value = txtTFSAC2.Text
        .Range("C" & r + 3).Value = txtTFSAC3.Text
        .Range("C" & r + 4).Value = txtTFSAC4.Text
        .Range("C" & r + 5).Value = txtTFSAC5.Text
        .Range("C" & r + 6).Value = IIf(optTFSACClassYes.Value = True, "Yes", "No")
        .Range("C" & r + 7).Value = IIf(optTFSACDivisionYes.Value = True, "Yes", "No")
        .Range("C" & r + 8).Value = txtTFSAC7.Text
        .Range("C" & r + 9).Value = txtTFSAC8.Text
        .Range("C" & r + 10).Value = txtTFSAC9.Text
        .Range("C" & r + 11).Value = txtTFSAC10.Text
    
        r = .Range("TFSAWithdrawals").Row
        .Range("C" & r + 1).Value = txtTFSAW1.Text
        .Range("C" & r + 2).Value = txtTFSAW2.Text
        .Range("C" & r + 3).Value = txtTFSAW3.Text
        .Range("C" & r + 4).Value = IIf(optTFSAWClassYes.Value = True, "Yes", "N0")
        .Range("C" & r + 5).Value = IIf(optTFSAWDivisionYes.Value = True, "Yes", "No")
    End With
End Sub

Private Sub CommandButton1_Click()
    With Sheet1
        .Range("HRClass").Value = "Class"
        i = .Range("HRClass").Row + 13
        j = .Range("HRDivision").Row
        .Range("C" & i - 12 & ":C" & i).ClearContents
        If i < j Then
            .Rows(i & ":" & j - 1).Delete
        End If
    End With
End Sub

Private Sub CommandButton2_Click()
    With Sheet1
        .Range("HRDivision").Value = "Division"
        i = .Range("HRDivision").Row + 13
        j = .Range("PlanSponsorStatement").Row - 2
        .Range("C" & i - 12 & ":C" & i).ClearContents
        If i < j Then
            .Rows(i & ":" & j).Delete
        End If
    End With
End Sub

Private Sub CommandButton3_Click()
    Dim strFile As String
    Dim invoiceRng As Range
    Dim pdfFile As String
    
    'Setting range to be printed
    Set invoiceRng = Range("A1:D550")
    'setting file name with a time stamp.
    strFile = "Checklist" & "_" & Format(Now(), "yyyymmdd_hhmmss") & ".pdf"
    'setting the name. The resultent pdf will be saved where the main file exists.
    pdfFile = ThisWorkbook.Path & Application.PathSeparator & strFile
    invoiceRng.ExportAsFixedFormat Type:=xlTypePDF, _
    Filename:=pdfFile, _
    Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, _
    IgnorePrintAreas:=True, _
    OpenAfterPublish:=True
End Sub

Private Sub CommandButton4_Click()
    Dim desktopPath As String
    Dim NewFile As String
    
    desktopPath = GetDesktop & "\"
    NewFile = desktopPath & Left$(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 4) & "_new.xlsm"
    ActiveSheet.Copy
    ActiveWorkbook.SaveAs NewFile, 52
    ActiveWorkbook.Close True
End Sub

Private Sub CommandButton5_Click()
    Dim lastRow As Integer
    
    lastRow = Sheet3.Range("TaxType").Row - 15
    Sheet3.Range("A1:A" & lastRow).EntireRow.Delete
    With Sheet1
        .Range("A1:A81").EntireRow.Copy
        Sheet3.Range("A1").Insert
        lastRow = 82
        If chkRRSP.Value = True Then
            i = .Range("RRSP").Row
            j = .Range("DPSP").Row - 2
            .Range("A" & i & ":A" & j).EntireRow.Copy
            Sheet3.Range("A" & lastRow).Insert
            lastRow = Sheet3.Range("B" & Sheet3.Range("TaxType").Row).End(xlUp).Row + 2
        End If
        If chkDPSP.Value = True Then
            i = .Range("DPSP").Row
            j = .Range("RPP").Row - 2
            .Range("A" & i & ":A" & j).EntireRow.Copy
            Sheet3.Range("A" & lastRow).Insert
            lastRow = Sheet3.Range("B" & Sheet3.Range("TaxType").Row).End(xlUp).Row + 2
        End If
        If chkRPP.Value = True Then
            i = .Range("RPP").Row
            j = .Range("TFSA").Row - 2
            .Range("A" & i & ":A" & j).EntireRow.Copy
            Sheet3.Range("A" & lastRow).Insert
            lastRow = Sheet3.Range("B" & Sheet3.Range("TaxType").Row).End(xlUp).Row + 2
        End If
        If chkTFSA.Value = True Then
            i = .Range("TFSA").Row
            j = .Range("SPP").Row - 2
            .Range("A" & i & ":A" & j).EntireRow.Copy
            Sheet3.Range("A" & lastRow).Insert
            lastRow = Sheet3.Range("B" & Sheet3.Range("TaxType").Row).End(xlUp).Row + 2
        End If
        If chkSPP.Value = True Then
            i = .Range("SPP").Row
            j = .Range("HRServices").Row - 3
            .Range("A" & i & ":A" & j).EntireRow.Copy
            Sheet3.Range("A" & lastRow).Insert
        End If
    End With
    
    Application.CutCopyMode = False
End Sub

Private Sub frmResetSheet_Click()
    Dim c As Range
    
    With Sheet1
        i = .Range("TFSAWithdrawalDivision").End(xlDown).Row
        .Range("C1:C" & i).ClearContents
        
        Set c = .UsedRange.Find(What:="Classes (if applicable)", LookIn:=xlValues)
        .Range(.Cells(c.Row + 2, 3), .Cells(c.Row + 2, 20)).ClearContents
    End With
    
    MsgBox "Sheet has been reset", vbInformation, "Success"
End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub optAdminstrator_Click()
    With Sheet1
        i = .Range("MemberInformation").Row + 2
        If optAdminstrator.Value = True Then .Range("C" & i).Value = optAdminstrator.Caption
    End With
End Sub

Private Sub optAffiliateNo_Click()
If optAffiliateNo.Value = True Then
[11:16].EntireRow.Hidden = True
End If
End Sub

Private Sub optAffiliateYes_Click()
If optAffiliateYes.Value = True Then
[11:16].EntireRow.Hidden = False
End If
End Sub

Private Sub optClassNo_Click()
If optClassNo.Value = True Then
Sheet1.Range("C34").Value = optClassNo.Caption & _
[35:36].EntireRow.Hidden = True
End If


End Sub

Private Sub optClassYes_Click()
If optClassYes.Value = True Then
Sheet1.Range("C34").Value = optClassYes.Caption
End If

End Sub

Private Sub optDivisionalNo_Click()
If optDivisionalNo.Value = True Then
[27:32].EntireRow.Hidden = True
End If

End Sub

Private Sub optDivisionalYes_Click()
If optDivisionalYes.Value = True Then
[27:32].EntireRow.Hidden = False
End If

End Sub

'==================================================================================================
'===============================     RRSP Eligibility Class & Division   ==========================
'RRSP Eligibility Class No
Private Sub optEClassNo_Click()
    With Sheet1
        i = .Range("RRSPEligibilityClass").Row
        j = .Range("RRSPEligibilityDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'RRSP Eligibility Class Yes
Private Sub optEClassYes_Click()
    With Sheet1
        i = .Range("RRSPEligibilityClass").Row
        j = .Range("RRSPEligibilityDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'RRSP Eligibility Division No
Private Sub optEDivisionNo_Click()
    With Sheet1
        i = .Range("RRSPEligibilityDivision").Row
        j = .Range("RRSPContribution").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'RRSP Eligibility Division Yes
Private Sub optEDivisionYes_Click()
    With Sheet1
        i = .Range("RRSPEligibilityDivision").Row
        j = .Range("RRSPContribution").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'==================================================================================================
'===============================     RRSP Contribution Class & Division  ==========================
'RRSP Contribution Class No
Private Sub optCClassNo_Click()
    With Sheet1
        i = .Range("RRSPContributionClass").Row
        j = .Range("RRSPContributionDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'RRSP Contribution Class Yes
Private Sub optCClassYes_Click()
    With Sheet1
        i = .Range("RRSPContributionClass").Row
        j = .Range("RRSPContributionDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'RRSP Contribution Division No
Private Sub optCDivisionNo_Click()
    With Sheet1
        a = .Range("RRSPContributionDivision").Row
        b = .Range("RRSPWithdrawals").Row - 1
        .Range(a & ":" & b).EntireRow.Hidden = True
    End With
End Sub

'RRSP Contribution Division Yes
Private Sub optCDivisionYes_Click()
    With Sheet1
        a = .Range("RRSPContributionDivision").Row
        b = .Range("RRSPWithdrawals").Row - 1
        .Range(a & ":" & b).EntireRow.Hidden = False
    End With
End Sub

Private Sub optManulife_Click()
    With Sheet1
        i = .Range("MemberInformation").Row + 1
        If optManulife.Value = True Then .Range("C" & i).Value = optManulife.Caption
    End With
End Sub

Private Sub optPersonnel_Click()
    With Sheet1
        i = .Range("MemberInformation").Row + 2
        If optPersonnel.Value = True Then .Range("C" & i).Value = optPersonnel.Caption
    End With
End Sub

Private Sub optPlanNo_Click()
If optPlanNo.Value = True Then
Sheet1.Range("C21").Value = optPlanNo.Caption
End If
End Sub

Private Sub optPlanYes_Click()
If optPlanYes.Value = True Then
Sheet1.Range("C21").Value = optPlanYes.Caption
End If
End Sub

Private Sub optSecNo_Click()
If optSecNo.Value = True Then
[23:26].EntireRow.Hidden = True
End If

End Sub

Private Sub optSecYes_Click()
If optSecYes.Value = True Then
[23:26].EntireRow.Hidden = False
End If
End Sub

Private Sub optSponsor_Click()
    With Sheet1
        i = .Range("MemberInformation").Row + 1
        If optSponsor.Value = True Then .Range("C" & i).Value = optSponsor.Caption
    End With
End Sub

'==================================================================================================
'================================     RRSP Withdrawal Class & Division  ===========================
'RRSP Withdrawal Class No
Private Sub optWClassNo_Click()
    With Sheet1
        i = .Range("RRSPWithdrawalClass").Row
        j = .Range("RRSPWithdrawalDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'RRSP Withdrawal Class Yes
Private Sub optWClassYes_Click()
    With Sheet1
        i = .Range("RRSPWithdrawalClass").Row
        j = .Range("RRSPWithdrawalDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'RRSP Withdrawal Division No
Private Sub optWDivisionNo_Click()
    With Sheet1
        i = .Range("RRSPWithdrawalDivision").Row
        j = .Range("DPSP").Row - 3
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'RRSP Withdrawal Division Yes
Private Sub optWDivisionYes_Click()
    With Sheet1
        i = .Range("RRSPWithdrawalDivision").Row
        j = .Range("DPSP").Row - 3
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'==================================================================================================
'==============================     DPSP Eligibility Class  & Division ============================
'DPSP Eligibility Class No
Private Sub optDPSPEClassNo_Click()
    With Sheet1
        i = .Range("DPSPEligibilityClass").Row
        j = .Range("DPSPEligibilityDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'DPSP Eligibility Class Yes
Private Sub optDPSPEClassYes_Click()
    With Sheet1
        i = .Range("DPSPEligibilityClass").Row
        j = .Range("DPSPEligibilityDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'DPSP Eligibility Division No
Private Sub optDPSPEDivisionNo_Click()
    With Sheet1
        i = .Range("DPSPEligibilityDivision").Row
        j = .Range("DPSPContribution").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'DPSP Eligibility Division Yes
Private Sub optDPSPEDivisionYes_Click()
    With Sheet1
        i = .Range("DPSPEligibilityDivision").Row
        j = .Range("DPSPContribution").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'==================================================================================================
'===============================     DPSP Contribution Class & Division  ==========================
'DPSP Contribution Class No
Private Sub optDPSPCClassNo_Click()
    With Sheet1
        i = .Range("DPSPContributionClass").Row
        j = .Range("DPSPContributionDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'DPSP Contribution Class Yes
Private Sub optDPSPCClassYes_Click()
    With Sheet1
        i = .Range("DPSPContributionClass").Row
        j = .Range("DPSPContributionDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'DPSP Contribution Division No
Private Sub optDPSPCDivisionNo_Click()
    With Sheet1
        i = .Range("DPSPContributionDivision").Row
        j = .Range("DPSPWithdrawals").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'DPSP Contribution Division Yes
Private Sub optDPSPCDivisionYes_Click()
    With Sheet1
        i = .Range("DPSPContributionDivision").Row
        j = .Range("DPSPWithdrawals").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'==================================================================================================
'================================     DPSP Withdrawal Class & Division  ===========================
'DPSP Withdrawal Class No
Private Sub optDPSPWClassNo_Click()
    With Sheet1
        i = .Range("DPSPWithdrawalClass").Row
        j = .Range("DPSPWithdrawalDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'DPSP Withdrawal Class Yes
Private Sub optDPSPWClassYes_Click()
    With Sheet1
        i = .Range("DPSPWithdrawalClass").Row
        j = .Range("DPSPWithdrawalDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'DPSP Withdrawal Division No
Private Sub optDPSPWDivisionNo_Click()
    With Sheet1
        i = .Range("DPSPWithdrawalDivision").Row
        j = .Range("RPP").Row - 11
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'DPSP Withdrawal Division Yes
Private Sub optDPSPWDivisionYes_Click()
    With Sheet1
        i = .Range("DPSPWithdrawalDivision").Row
        j = .Range("RPP").Row - 11
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'==================================================================================================
'==============================     RPP Eligibility Class  & Division =============================
'RPP Eligibility Class No
Private Sub optRPPEClassNo_Click()
    With Sheet1
        i = .Range("RPPEligibilityClass").Row
        j = .Range("RPPEligibilityDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'RPP Eligibility Class Yes
Private Sub optRPPEClassYes_Click()
    With Sheet1
        i = .Range("RPPEligibilityClass").Row
        j = .Range("RPPEligibilityDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'RPP Eligibility Division No
Private Sub optRPPEDivisionNo_Click()
    With Sheet1
        i = .Range("RPPEligibilityDivision").Row
        j = .Range("RPPContribution").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'RPP Eligibility Division Yes
Private Sub optRPPEDivisionYes_Click()
    With Sheet1
        i = .Range("RPPEligibilityDivision").Row
        j = .Range("RPPContribution").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'==================================================================================================
'===============================     RPP Contribution Class & Division  ===========================
'RPP Contribution Class No
Private Sub optRPPCClassNo_Click()
    With Sheet1
        i = .Range("RPPContributionClass").Row
        j = .Range("RPPContributionDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'RPP Contribution Class Yes
Private Sub optRPPCClassYes_Click()
    With Sheet1
        i = .Range("RPPContributionClass").Row
        j = .Range("RPPContributionDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'RPP Contribution Division No
Private Sub optRPPCDivisionNo_Click()
    With Sheet1
        i = .Range("RPPContributionDivision").Row
        j = .Range("RPPWithdrawals").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'RPP Contribution Division Yes
Private Sub optRPPCDivisionYes_Click()
    With Sheet1
        i = .Range("RPPContributionDivision").Row
        j = .Range("RPPWithdrawals").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'==================================================================================================
'================================     RPP Withdrawal Class & Division  ============================
'RPP Withdrawal Class No
Private Sub optRPPWClassNo_Click()
    With Sheet1
        i = .Range("RPPWithdrawalClass").Row
        j = .Range("RPPWithdrawalDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'RPP Withdrawal Class Yes
Private Sub optRPPWClassYes_Click()
    With Sheet1
        i = .Range("RPPWithdrawalClass").Row
        j = .Range("RPPWithdrawalDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'RPP Withdrawal Division No
Private Sub optRPPWDivisionNo_Click()
    With Sheet1
        i = .Range("RPPWithdrawalDivision").Row
        j = .Range("TFSA").Row - 11
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'RPP Withdrawal Division Yes
Private Sub optRPPWDivisionYes_Click()
    With Sheet1
        i = .Range("RPPWithdrawalDivision").Row
        j = .Range("TFSA").Row - 11
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'==================================================================================================
'==============================     TFSA Eligibility Class  & Division ============================
'TFSA Eligibility Class No
Private Sub optTFSAEClassNo_Click()
    With Sheet1
        i = .Range("TFSAEligibilityClass").Row
        j = .Range("TFSAEligibilityDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'TFSA Eligibility Class Yes
Private Sub optTFSAEClassYes_Click()
    With Sheet1
        i = .Range("TFSAEligibilityClass").Row
        j = .Range("TFSAEligibilityDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'TFSA Eligibility Division No
Private Sub optTFSAEDivisionNo_Click()
    With Sheet1
        i = .Range("TFSAEligibilityDivision").Row
        j = .Range("TFSAContribution").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'TFSA Eligibility Division Yes
Private Sub optTFSAEDivisionYes_Click()
    With Sheet1
        i = .Range("TFSAEligibilityDivision").Row
        j = .Range("TFSAContribution").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'==================================================================================================
'==============================     TFSA Contrbution Class  & Division ============================
'TFSA Contribution Class No
Private Sub optTFSACClassNo_Click()
    With Sheet1
        i = .Range("TFSAContributionClass").Row
        j = .Range("TFSAContributionDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'TFSA Contribution Class Yes
Private Sub optTFSACClassYes_Click()
    With Sheet1
        i = .Range("TFSAContributionClass").Row
        j = .Range("TFSAContributionDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'TFSA Contribution Division No
Private Sub optTFSACDivisionNo_Click()
    With Sheet1
        i = .Range("TFSAContributionDivision").Row
        j = .Range("TFSAWithdrawals").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'TFSA Contribution Division Yes
Private Sub optTFSACDivisionYes_Click()
    With Sheet1
        i = .Range("TFSAContributionDivision").Row
        j = .Range("TFSAWithdrawals").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'==================================================================================================
'==============================     TFSA Withdrawal Class  & Division =============================
'TFSA Withdrawal Class No
Private Sub optTFSAWClassNo_Click()
    With Sheet1
        i = .Range("TFSAWithdrawalClass").Row
        j = .Range("TFSAWithdrawalDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'TFSA Withdrawal Class Yes
Private Sub optTFSAWClassYes_Click()
    With Sheet1
        i = .Range("TFSAWithdrawalClass").Row
        j = .Range("TFSAWithdrawalDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'TFSA Withdrawal Division No
Private Sub optTFSAWDivisionNo_Click()
    With Sheet1
        i = .Range("TFSAWithdrawalDivision").Row
        j = .Range("SPP").Row - 3
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'TFSA Withdrawal Division Yes
Private Sub optTFSAWDivisionYes_Click()
    With Sheet1
        i = .Range("TFSAWithdrawalDivision").Row
        j = .Range("SPP").Row - 3
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'==================================================================================================
'==============================     SPP Eligibility Class  & Division =============================
'SPP Eligibility Class No
Private Sub optSPPEClassNo_Click()
    With Sheet1
        i = .Range("SPPEligibilityClass").Row
        j = .Range("SPPEligibilityDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'SPP Eligibility Class Yes
Private Sub optSPPEClassYes_Click()
    With Sheet1
        i = .Range("SPPEligibilityClass").Row
        j = .Range("SPPEligibilityDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'SPP Eligibility Division No
Private Sub optSPPEDivisionNo_Click()
    With Sheet1
        i = .Range("SPPEligibilityDivision").Row
        j = .Range("SPPContribution").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'SPP Eligibility Division Yes
Private Sub optSPPEDivisionYes_Click()
    With Sheet1
        i = .Range("SPPEligibilityDivision").Row
        j = .Range("SPPContribution").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'==================================================================================================
'==============================     SPP Contribution Class  & Division ============================
'SPP Contribution Class No
Private Sub optSPPCClassNo_Click()
    With Sheet1
        i = .Range("SPPContributionClass").Row
        j = .Range("SPPContributionDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'SPP Contribution Class Yes
Private Sub optSPPCClassYes_Click()
    With Sheet1
        i = .Range("SPPContributionClass").Row
        j = .Range("SPPContributionDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'SPP Contribution Division NO
Private Sub optSPPCDivisionNo_Click()
    With Sheet1
        i = .Range("SPPContributionDivision").Row
        j = .Range("SPPWithdrawals").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'SPP Contribution Division Yes
Private Sub optSPPCDivisionYes_Click()
    With Sheet1
        i = .Range("SPPContributionDivision").Row
        j = .Range("SPPWithdrawals").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'==================================================================================================
'==============================     SPP Withdrawals Class  & Division =============================
'SPP Withdrawals Class No
Private Sub optSPPWClassNo_Click()
    With Sheet1
        i = .Range("SPPWithdrawalClass").Row
        j = .Range("SPPWithdrawalDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'SPP Withdrawals Class Yes
Private Sub optSPPWClassYes_Click()
    With Sheet1
        i = .Range("SPPWithdrawalClass").Row
        j = .Range("SPPWithdrawalDivision").Row - 1
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

'SPP Withdrawals Division No
Private Sub optSPPWDivisionNo_Click()
    With Sheet1
        i = .Range("SPPWithdrawalDivision").Row
        j = .Range("HRServices").Row - 9
        .Range(i & ":" & j).EntireRow.Hidden = True
    End With
End Sub

'SPP Withdrawals Division Yes
Private Sub optSPPWDivisionYes_Click()
    With Sheet1
        i = .Range("SPPWithdrawalDivision").Row
        j = .Range("HRServices").Row - 9
        .Range(i & ":" & j).EntireRow.Hidden = False
    End With
End Sub

Private Sub optHRNo_Click()
If optHRNo.Value = True Then
txtHR15.Locked = True
End If

End Sub

Private Sub UserForm_Activate()
    '===========================================================================================
    '=================================  Company Information ====================================
    '===========================================================================================
    cmbOrgType.AddItem ("Corporation")
    cmbOrgType.AddItem ("Union")
    cmbOrgType.AddItem ("Non-Profit Organization")
    cmbOrgType.AddItem ("Sole Proprietorship")
    cmbOrgType.AddItem ("Club")
    cmbOrgType.AddItem ("Society")
    cmbOrgType.AddItem ("Association")
    cmbOrgType.AddItem ("Partnership")
    cmbOrgType.AddItem ("Public Body")
    cmbOrgType.AddItem ("Other")
    
    cmbFreqContrib.AddItem ("Weekly")
    cmbFreqContrib.AddItem ("Bi-Weekly")
    cmbFreqContrib.AddItem ("Monthly")
    cmbFreqContrib.AddItem ("Other")
    
    cmbContribType.AddItem ("Pad")
    cmbContribType.AddItem ("Wire")
    cmbContribType.AddItem ("Cheque")
    cmbContribType.AddItem ("Other")
    
    MultiPage1.Value = 0
    
    Dim c As Range
    Dim i As Integer
    
    With Sheet1
        'add eligibility class combobox
        Set c = .UsedRange.Find(What:="Classes (if applicable)", LookIn:=xlValues)
        
        If Not c Is Nothing Then
            cboEClass.Clear
            If Len(.Range("C" & c.Row + 2).Value) > 0 Then
                i = 3
                While Len(.Cells(c.Row + 2, i).Value) > 0
                    cboEClass.AddItem .Cells(c.Row + 2, i).Value
                    i = i + 1
                Wend
            End If
        End If
        
        'add eligibility division combobox
        Set c = .UsedRange.Find(What:="Divisional (if applicable)", LookIn:=xlValues)
        
        If Not c Is Nothing Then
            cboEDivision.Clear
            If .Range("Z" & c.Row).Value = 1 Then
                If Len(.Range("C" & c.Row + 1).Value) > 0 And .Range("C" & c.Row + 1) <> "N/A" Then
                    cboEDivision.AddItem .Range("C" & c.Row + 1).Value
                End If
            ElseIf .Range("Z" & c.Row).Value > 1 Then
                j = 1
                For i = 1 To .Range("Z" & c.Row).Value
                    cboEDivision.AddItem .Range("C" & c.Row + j).Value
                    j = j + 6
                Next i
            End If
        End If
        
        'check affiliate section and adjust option button
        Set c = .UsedRange.Find(What:="Affiliated company Information", LookIn:=xlValues)
        If Not c Is Nothing Then optAffiliateYes.Value = True Else optAffiliateNo.Value = True
        
        'check division section and adjust option button
        Set c = .UsedRange.Find(What:="Divisional (if applicable)", LookIn:=xlValues)
        If Not c Is Nothing Then
            If .Range("C" & c.Row + 5).Value = "Yes" Then optDivisionalYes.Value = True Else optDivisionalNo.Value = True
        Else
            optDivisionalNo.Value = True
        End If
        
        'check class section and adjust option button
        Set c = .UsedRange.Find(What:="Classes (if applicable)", LookIn:=xlValues)
        If Not c Is Nothing Then
            If .Range("C" & c.Row + 1).Value = "Yes" Then optClassYes.Value = True Else optClassNo.Value = True
        Else
            optClassNo.Value = True
        End If
        
        'check secondary section and adjust option button
        Set c = .UsedRange.Find(What:="Is a Secondary Contact required?", LookIn:=xlValues)
        If Not c Is Nothing Then
            If .Range("C" & c.Row).Value = "Yes" Then optSecYes.Value = True Else optSecNo.Value = True
        End If
        
        'check appropriate check boxes
        Set c = .UsedRange.Find(What:="Plan Type: RRSP", LookIn:=xlValues)
        chkRRSP.Value = Not c Is Nothing
        
        Set c = .UsedRange.Find(What:="Plan Type: DPSP", LookIn:=xlValues)
        chkDPSP.Value = Not c Is Nothing
        
        Set c = .UsedRange.Find(What:="Plan Type: RPP", LookIn:=xlValues)
        chkRPP.Value = Not c Is Nothing
        
        Set c = .UsedRange.Find(What:="Plan Type: TFSA", LookIn:=xlValues)
        chkTFSA.Value = Not c Is Nothing
        
        Set c = .UsedRange.Find(What:="PLan Type: SPP", LookIn:=xlValues)
        chkSPP.Value = Not c Is Nothing
    End With
End Sub

Function GetDesktop() As String
    Dim oWSHShell As Object

    Set oWSHShell = CreateObject("WScript.Shell")
    GetDesktop = oWSHShell.SpecialFolders("Desktop")
    Set oWSHShell = Nothing
End Function
