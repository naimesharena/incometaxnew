Attribute VB_Name = "Pincodechanges"

' Function state_Validation(ByVal pin_trgt_adrs As String, ByVal state_trgt_adrs As String) As Boolean
'    state_Validation = True
'
'    Dim PinCode As Range
'
'            If Range(pin_trgt_adrs).Value <> "" Then
'              Dim State1 As String
'                State1 = Application.VLookup(Range(pin_trgt_adrs).Value, Sheet5.Range("All_Pincode_V"), 2, False)
'                If Not (State1 = UCase(Mid(Range(state_trgt_adrs).Value, 4)) Or Sheet1.StateMatchesPin(Range(state_trgt_adrs).Value, Range(pin_trgt_adrs).Value)) Then
'                    state_Validation = False
'
'                    'Range(pin_trgt_adrs).value = ""
'                End If
'            End If
'
'
'End Function
 
                 
'Function PinState_codes_validation(sheetname1 As Worksheet, ByVal pin_trgt_adrs As String, ByVal state_trgt_adrs As String, ByVal country_trgt_adrs As String, ByVal Zip_trgt_adrs As String)
'
'Dim pincode1 As String
'DupFlag = False
'pincode1 = Application.IsError(Application.VLookup(sheetname1.Range(pin_trgt_adrs).Value, Sheet5.Range("Duplicate_Pincode_List"), 1, False))
'If pincode1 = "False" Then DupFlag = True
'
'If DupFlag = True Then
'   If sheetname1.Range(state_trgt_adrs).Value <> "" Then
'       If Not Sheet1.StateMatchesPin(Range(state_trgt_adrs).Value, sheetname1.Range(pin_trgt_adrs).Value) Then
'           sheetname1.Range(state_trgt_adrs).Value = "(Select)"
'       End If
'   End If
'End If
                 
'If DupFlag = False Then
'   Dim StateName
'   Dim StateCode
'
'   StateName = Application.VLookup(sheetname1.Range(pin_trgt_adrs).Value, Sheet5.Range("All_Pincode_V"), 2, False)
'   If IsError(StateName) Then
'       If sheetname1.Range(pin_trgt_adrs).Value <> "" Then
'
'           MsgBox ("Invalid Pincode")
'           sheetname1.Range(pin_trgt_adrs).ClearContents
'       End If
'   Else
'       StateCode = Application.VLookup(StateName, Sheet5.Range("CS:CT"), 2, False)
'       sheetname1.Range(state_trgt_adrs).Value = StateCode
'       sheetname1.Range(country_trgt_adrs).Value = "91-INDIA"
'       sheetname1.Unprotect Password:=mdCommon.getmsgstate
'   End If
'End If
'
'
'End Function


 

