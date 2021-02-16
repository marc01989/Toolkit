VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmERRDispute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Compare Database
Dim strComment As Variant
Dim varOpen As Variant
Dim varCaseID As Variant
Dim varFormName As Variant

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'this is tied to frmHICsUpdate and frmHICSUpdate2
'these generate the disputes for HICS cases
'form cannot be opened without passing case id upon Form_Load()
'-----------------------------------------------------------------------------------------------------------------------------------------------------------


Private Sub Form_Load()
    Dim strArgs As String
    Dim varArgs As Variant
    
    'passes caseid and form name from frmHICSUPdate or frmHICSUPdate2
    strArgs = Forms!frmERRDispute.OpenArgs & vbNullString
        If Not IsNull(strArgs) And InStr(strArgs, "|") <> 0 Then
        varArgs = Split(strArgs, "|")
        varCaseID = varArgs(0)
        varFormName = varArgs(1)
        ResetForm
        Me.AutoResize = True
        Me.FitToScreen = True
        DoCmd.Maximize
    Else
        MsgBox ("Cannot open dispute form without valid HICS case")
        DoCmd.Close acForm, "frmERRDispute", acSaveNo
    End If
    
End Sub
'makes appropriate textboxes visible and populates txtComment with the statement that goes on the dispute
Private Sub cboERRDisputeType_AfterUpdate()
    varOpen = cboERRDisputeType.Value
    ResetForm
    SelectComment
End Sub

Private Sub cmdSubmitERR_Click()

'checks that boxes are not null
    Select Case varOpen
        Case 1:
            If Nz(txtMemberName.Value, "") = "" Or Nz(txtAPTC.Value, "") = "" Or Nz(txtPolicyID.Value, "") = "" Or Nz(txtSBSID.Value, "") = "" Or Nz(txtStartDate.Value, "") = "" Then
                MsgBox "You must supply values for all relevant fields", vbOKOnly
                Exit Sub
            End If
        Case 2:
           If Nz(txtMemberName.Value, "") = "" Or Nz(txtTotalPremAmt.Value, "") = "" Or Nz(txtPolicyID.Value, "") = "" Or Nz(txtSBSID.Value, "") = "" Or Nz(txtStartDate.Value, "") = "" Then
                MsgBox "You must supply values for all relevant fields", vbOKOnly
                Exit Sub
          End If
            
        Case 3:
            If Nz(txtMemberName.Value, "") = "" Or Nz(txtStartDate.Value, "") = "" Or Nz(txtEndDate.Value, "") = "" Or Nz(txtPolicyID.Value, "") = "" Or Nz(txtSBSID.Value, "") = "" Then
                MsgBox "You must supply values for all relevant fields", vbOKOnly
                Exit Sub
            End If
                      
        Case 4:
            If Nz(txtMemberName.Value, "") = "" Or Nz(txtEnrollmentBlockerCode.Value, "") = "" Or Nz(txtStartDate.Value, "") = "" Or Nz(txtTotalPremAmt.Value, "") = "" Or Nz(txtAPTC.Value, "") = "" Or Nz(txtCSR.Value, "") = "" Or Nz(txtSBSID.Value, "") = "" Or Nz(txtPolicyID.Value, "") = "" Or Nz(txtNewStartDate.Value, "") = "" Then
              MsgBox "You must supply values for all relevant fields", vbOKOnly
              Exit Sub
            End If
        Case 5:
          If Nz(txtMemberName.Value, "") = "" Or Nz(txtQHPID.Value, "") = "" Or Nz(txtSBSID.Value, "") = "" Or Nz(txtStartDate.Value, "") = "" Or Nz(txtAPTC.Value, "") = "" Or Nz(txtPolicyID.Value, "") = "" Then
            MsgBox "You must supply values for all relevant fields", vbOKOnly
            Exit Sub
        End If
        Case 6:
            If Nz(txtAPTC.Value, "") = "" Or Nz(txtSBSID.Value, "") = "" Or Nz(txtRemovedMember.Value, "") = "" Or Nz(txtStartDate.Value, "") = "" Or Nz(txtPolicyID.Value, "") = "" Or Nz(txtTotalPremAmt.Value, "") = "" Or Nz(txtSBSRemaining.Value, "") = "" Or Nz(txtRemovedMember.Value, "") = "" Then
                MsgBox "You must supply values for all relevant fields", vbOKOnly
                Exit Sub
            End If
        Case 7:
            If Nz(txtStartDate.Value, "") = "" Or Nz(txtMemberName.Value, "") = "" Or Nz(txtSBSID.Value, "") = "" Or Nz(txtSBSCurrent.Value, "") = "" Or Nz(txtPolicyID.Value, "") = "" Then
                MsgBox "You must supply values for all relevant fields", vbOKOnly
                Exit Sub
            End If
    End Select

'if passed validation, prompt user if they want to add to frmHICSUpdate
    If MsgBox("Apply ERR to HICS case " & varCaseID & "?", vbYesNo + vbQuestion, "Submit ERR?") = vbYes Then
        Forms(varFormName).txtERR = "Y"
        Forms(varFormName).txtERRNoteReview = strComment
        Call Forms(varFormName).txtERRNoteReview_Change
        DoCmd.Close acForm, "frmERRDispute", acSaveNo
    Else
        Exit Sub
    End If
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmHICSUpdate; SubmitCaseWorkNote()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select

    

End Sub



'-------------------------------------------------------------------------------------------------------------------------------------------------------------
' after txtboxes updated below, refresh the string in txtComment box
'-------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub txtAPTC_AfterUpdate()
    SelectComment
End Sub

Private Sub txtCSR_AfterUpdate()
    SelectComment
End Sub

Private Sub txtEffectiveDate_AfterUpdate()
    SelectComment
End Sub

Private Sub txtEndDate_AfterUpdate()
    SelectComment
End Sub

Private Sub txtEnrollmentBlockerCode_AfterUpdate()
    SelectComment
End Sub

Private Sub txtMemberName_AfterUpdate()
    SelectComment
End Sub

Private Sub txtNewStartDate_AfterUpdate()
    SelectComment
End Sub

Private Sub txtPolicyID_AfterUpdate()
    SelectComment
End Sub

Private Sub txtQHPID_AfterUpdate()
    SelectComment
End Sub

Private Sub txtRemovedMember_AfterUpdate()
    SelectComment
End Sub

Private Sub txtSBSID_AfterUpdate()
    SelectComment
End Sub

Private Sub txtSBSRemaining_AfterUpdate()
    SelectComment
End Sub

Private Sub txtStartDate_AfterUpdate()
    SelectComment
End Sub

Private Sub txtTotalPremAmt_AfterUpdate()
    SelectComment
End Sub
Private Sub txtSBSNewAssigned_AfterUpdate()
    SelectComment
End Sub
Private Sub txtSBSCurrent_AfterUpdate()
    SelectComment
End Sub

'-------------------------------------------------------------------------------------------------------------------------------------------------------------
'sets comment string to the values of the textboxes, refreshes the value in txtComment
'makes appropriate
'-------------------------------------------------------------------------------------------------------------------------------------------------------------

Private Sub SelectComment()
      Select Case varOpen
    Case 1: strComment = _
        "ER&R: Please update the APTC to $" & txtAPTC.Value & " for Policy ID " & txtPolicyID.Value & ", belonging to " & txtMemberName.Value & ", Subscriber " & txtSBSID.Value & " effective: " & txtStartDate.Value & ". Please see narrative and additional casework notes."
          txtMemberName.Visible = True
          txtAPTC.Visible = True
          txtPolicyID.Visible = True
          txtSBSID.Visible = True
          txtStartDate.Visible = True
     Case 2: strComment = _
        "ER&R: Please update the total premium amount to $" & txtTotalPremAmt.Value & " for policy ID " & txtPolicyID.Value & " belonging to " & txtMemberName.Value & ", subscriber " & txtSBSID.Value & " effective " & txtStartDate & ". Please see narrative and additional casework notes."
          txtMemberName.Visible = True
          txtTotalPremAmt.Visible = True
          txtPolicyID.Visible = True
          txtSBSID.Visible = True
          txtStartDate.Visible = True
      Case 3: strComment = _
         "ER&R: Policy ID " & txtPolicyID.Value & " was terminated NLE and needs to be updated. The start date is " & txtStartDate.Value & " and the end date is " & txtEndDate.Value & " belonging to " & txtMemberName.Value & ", subscriber " & txtSBSID.Value & ". Please see narrative and additional casework notes."
           txtMemberName.Visible = True
           txtStartDate.Visible = True
           txtEndDate.Visible = True
           txtPolicyID.Visible = True
            txtSBSID.Visible = True
      Case 4: strComment = _
         "ER&R: Consumer encountered Enrollment Blocker " & txtEnrollmentBlockerCode.Value & " original benefit start date is " & txtStartDate.Value & ", total premium $" & txtTotalPremAmt.Value & ", APTC $" & txtAPTC.Value & _
         ", CSR variant " & txtCSR.Value & " for policy ID " & txtPolicyID.Value & " belonging to " & txtMemberName.Value & ", subscriber " & txtSBSID.Value & ". New benefit start date is " & txtNewStartDate.Value & ". Please see narrative and additional casework notes."
           txtMemberName.Visible = True
           txtEnrollmentBlockerCode.Visible = True
           txtStartDate.Visible = True
           txtTotalPremAmt.Visible = True
           txtAPTC.Visible = True
            txtCSR.Visible = True
            txtSBSID.Visible = True
            txtPolicyID.Visible = True
            txtNewStartDate.Visible = True
            
        Case 5: strComment = _
        "ER&R: Please update the QHP ID/Variant to " & txtQHPID.Value & " for policy ID " & txtPolicyID.Value & " belonging to " & txtMemberName.Value & ", subscriber " & txtSBSID.Value & ". The effective date is now " & txtStartDate.Value & " and the APTC is $" & txtAPTC.Value & ". Please see narrative and additional casework notes."
           txtMemberName.Visible = True
           txtQHPID.Visible = True
            txtPolicyID.Visible = True
            txtSBSID.Visible = True
            txtStartDate.Visible = True
            txtAPTC.Visible = True

        Case 6: strComment = _
        "ER&R: " & txtRemovedMember.Value & " should be removed from policy ID " & txtPolicyID.Value & " effective " & txtStartDate.Value & ". " & txtSBSRemaining.Value & ", subscriber " & txtSBSID.Value & " will remain the subscriber. The total premium without " & txtRemovedMember.Value & " is $" & txtTotalPremAmt.Value & " and the APTC is $" & txtAPTC.Value & " effective " & txtStartDate & ". Please see narrative and additional casework notes."
            txtRemovedMember.Visible = True
            txtStartDate.Visible = True
            txtPolicyID.Visible = True
            txtTotalPremAmt.Visible = True
            txtSBSRemaining.Visible = True
            txtSBSID.Visible = True
            txtAPTC.Visible = True
        Case 7: strComment = _
        "ER&R: " & txtMemberName.Value & ", exchange ID " & txtSBSID.Value & " should be made the new subscriber for policy ID " & txtPolicyID.Value & " effective " & txtStartDate & ". " & txtSBSCurrent & " is the current subscriber. Please see narrative and additional casework notes."
            txtPolicyID.Visible = True
            txtSBSCurrent.Visible = True
            txtSBSID.Visible = True
            txtMemberName.Visible = True
            txtStartDate.Visible = True
        End Select

        txtComment = strComment
End Sub

'resets all textboxes to ""
Private Sub ResetForm()
    Dim ctrl As Control
        For Each ctrl In Controls
            If TypeOf ctrl Is TextBox And ctrl.Name <> "txtComment" Then
                ctrl.Visible = False
                ctrl.Value = ""
                txtComment.Visible = True
            End If
        Next ctrl

            
End Sub
