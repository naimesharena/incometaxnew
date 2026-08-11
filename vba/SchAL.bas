Attribute VB_Name = "SchAL"
'Option Explicit
'
'
'Dim msgboxAL As String
'Public TotalIncome_8b As Variant
'Public ImmovableAssetLand_AL As Variant
'Public ImmovableAssetBuilding_AL As Variant
'Public DepositsInBank_AL As Variant
'Public SharesAndSecurities_AL As Variant
'Public InsurancePolicies_AL As Variant
'Public LoansAndAdvancesGiven_AL As Variant
'Public LiabilityInRelatAssets_AL As Variant
'Public VehiclYachtsBoatsAircrafts_AL As Variant
'Public ArchCollDrawPaintSulpArt_AL As Variant
'Public JewelleryBullionEtc_AL As Variant
'Public TotalImmovableAssets As Variant
'Public CashInHand_AL As Variant
'
'Sub Cmd_Validate_AL_Click()
'SchAL.ValidateSchAL
'MsgBox ("Schedule AL is ok")
'CloseMsg
'End Sub
'
'Sub NextAL_Click()
'Dim sourceSheet As Worksheet
'    Set sourceSheet = ThisWorkbook.Sheets("80G")
'    sourceSheet.Activate
'End Sub
'
'Sub PrevAL_Click()
'Dim a As Worksheet
'    Set a = ThisWorkbook.Sheets("Taxes Paid and Verification")
'    a.Activate
'End Sub
'
'Sub ValidateSchAL()
'    If Not ValidatesheetSchAL Then
'        'Sheet9.Activate
'        'MsgBox msgboxAL, vbOKOnly, "ITR-1"
'        fmsgbox (msgboxAL)
'        CloseMsg
'    End If
'End Sub
'Function ValidatesheetSchAL() As Boolean
'    ValidatesheetSchAL = True
'
'    msgboxAL = ""
'
'    TotalIncome_8b = Sheet1.Range("IncD.TotalIncome").Value
'
'    ImmovableAssetLand_AL = Sheet9.Range("AL.ImmovableAssetLand").Value
'    ImmovableAssetBuilding_AL = Sheet9.Range("AL.ImmovableAssetBuilding").Value
'
'    'DepositsInBank_AL = Sheet9.Range("AL.DepositsInBank").Value
'    'SharesAndSecurities_AL = Sheet9.Range("AL.SharesAndSecurities").Value
'    'InsurancePolicies_AL = Sheet9.Range("AL.InsurancePolicies").Value
'    'LoansAndAdvancesGiven_AL = Sheet9.Range("AL.LoansAndAdvancesGiven").Value
'    CashInHand_AL = Sheet9.Range("AL.CashInHand").Value
'
'    JewelleryBullionEtc_AL = Sheet9.Range("AL.JewelleryBullionEtc").Value
'    'ArchCollDrawPaintSulpArt_AL = Sheet9.Range("AL.ArchCollDrawPaintSulpArt").Value
'    VehiclYachtsBoatsAircrafts_AL = Sheet9.Range("AL.VehiclYachtsBoatsAircrafts").Value
'    TotalImmovableAssets = Sheet9.Range("al.TotalImmovablMovablAssets").Value
'    LiabilityInRelatAssets_AL = Sheet9.Range("AL.LiabilityInRelatAssets").Value
'
'    If ImmovableAssetLand_AL < 0 Then
'        msgboxAL = msgboxAL + "* Amount in Immovable Asset(Land) should be Numeric, Non Negative, not exceeding 14 digits " & Chr(13)
'        ValidatesheetSchAL = False
'    End If
'
'    If ImmovableAssetBuilding_AL < 0 Then
'        msgboxAL = msgboxAL + "* Amount in Immovable Asset(Building) should be Numeric, Non Negative, not exceeding 14 digits " & Chr(13)
'        ValidatesheetSchAL = False
'    End If
'
'    If CashInHand_AL < 0 Then
'        msgboxAL = msgboxAL + "* Amount in Movable Asset(Cash in hand) should be Numeric, Non Negative, not exceeding 14 digits " & Chr(13)
'        ValidatesheetSchAL = False
'    End If
'
'    If JewelleryBullionEtc_AL < 0 Then
'        msgboxAL = msgboxAL + "* Amount in Movable Asset(Jewellery,bullion etc.) should be Numeric, Non Negative, not exceeding 14 digits " & Chr(13)
'        ValidatesheetSchAL = False
'    End If
'
'    If VehiclYachtsBoatsAircrafts_AL < 0 Then
'        msgboxAL = msgboxAL + "* Amount in Movable Asset(Vehicles,yachts,boats and aircrafts) should be Numeric, Non Negative, not exceeding 14 digits " & Chr(13)
'        ValidatesheetSchAL = False
'    End If
'
'    If LiabilityInRelatAssets_AL < 0 Then
'        msgboxAL = msgboxAL + "* Amount in Liability in relation to Assets at A should be Numeric, Non Negative, not exceeding 14 digits " & Chr(13)
'        ValidatesheetSchAL = False
'    End If
'
'    If Len(TotalImmovableAssets) > 14 Then
'            msgboxAL = msgboxAL + "* Total Immovable asset cannot be more than 14 digits " & Chr(13)
'            ValidatesheetSchAL = False
'    End If
'
'    If TotalIncome_8b > 5000000 Then
'        noOfProcessSub = 5
'        subProcCaption = "Validating AL"
'
'        If ImmovableAssetLand_AL = "" Or IsEmpty(ImmovableAssetLand_AL) Then
'            msgboxAL = msgboxAL + "* Land in Schedule Asset and Liability in Sheet : AL  is Mandatory" & Chr(13)
'            ValidatesheetSchAL = False
'        End If
'        UpdateProgressBar
'
'        If ImmovableAssetBuilding_AL = "" Or IsEmpty(ImmovableAssetBuilding_AL) Then
'            msgboxAL = msgboxAL + "* Building in Schedule Asset and Liability in Sheet : AL  is Mandatory" & Chr(13)
'            ValidatesheetSchAL = False
'        End If
'
'        'If DepositsInBank_AL = "" Or IsEmpty(DepositsInBank_AL) Then
'             'msgboxAL = msgboxAL + "Deposits in bank in Schedule Asset and Liability in Sheet : AL  is Mandatory" & Chr(13)
'            'ValidatesheetSchAL = False
'        'End If
'
'        UpdateProgressBar
'        'If SharesAndSecurities_AL = "" Or IsEmpty(SharesAndSecurities_AL) Then
'            'msgboxAL = msgboxAL + "Shares and Securities in Schedule Asset and Liability in Sheet : AL  is Mandatory" & Chr(13)
'            'ValidatesheetSchAL = False
'        'End If
'
'        'If InsurancePolicies_AL = "" Or IsEmpty(InsurancePolicies_AL) Then
'             'msgboxAL = msgboxAL + "Insurance Policies in Schedule Asset and Liability in Sheet : AL  is Mandatory" & Chr(13)
'            'ValidatesheetSchAL = False
'        'End If
'
'        'If LoansAndAdvancesGiven_AL = "" Or IsEmpty(LoansAndAdvancesGiven_AL) Then
'             'msgboxAL = msgboxAL + "Loans and Advances Given in Schedule Asset and Liability in Sheet : AL  is Mandatory" & Chr(13)
'            'ValidatesheetSchAL = False
'        'End If
'        UpdateProgressBar
'        If CashInHand_AL = "" Or IsEmpty(CashInHand_AL) Then
'            msgboxAL = msgboxAL + "* Cash in hand  in Schedule Asset and Liability in Sheet : AL  is Mandatory" & Chr(13)
'            ValidatesheetSchAL = False
'        End If
'
'        If JewelleryBullionEtc_AL = "" Or IsEmpty(JewelleryBullionEtc_AL) Then
'            msgboxAL = msgboxAL + "* Jewellery / Bullion etc in Schedule Asset and Liability in Sheet : AL  is Mandatory" & Chr(13)
'            ValidatesheetSchAL = False
'        End If
'
'        UpdateProgressBar
'        'If ArchCollDrawPaintSulpArt_AL = "" Or IsEmpty(ArchCollDrawPaintSulpArt_AL) Then
'            ' msgboxAL = msgboxAL + "Arch Collections etc in Schedule Asset and Liability in Sheet : AL  is Mandatory" & Chr(13)
'            'ValidatesheetSchAL = False
'        'End If
'
'        If VehiclYachtsBoatsAircrafts_AL = "" Or IsEmpty(VehiclYachtsBoatsAircrafts_AL) Then
'            msgboxAL = msgboxAL + "* Vehicles, Yachts etc in Schedule Asset and Liability in Sheet : AL  is Mandatory" & Chr(13)
'            ValidatesheetSchAL = False
'        End If
'        UpdateProgressBar
'
'        If LiabilityInRelatAssets_AL = "" Or IsEmpty(LiabilityInRelatAssets_AL) Then
'            msgboxAL = msgboxAL + "* Liability in Schedule Asset and Liability in Sheet : AL  is Mandatory" & Chr(13)
'            ValidatesheetSchAL = False
'        End If
'    End If
'End Function
