VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmHICSUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim db As Database
Dim rs As Recordset
Dim strUser As String, strInteraction As String, strCaseType As String, strCaseAction As String
Dim bolCaseResolved As Boolean
Private Declare Function apiGetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'REFERENCE REQUIRED - Microsoft VBScript Regular Expression 5.5
    'https://regex101.com/
    
'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'this is the main form where hics cases are worked
'-----------------------------------------------------------------------------------------------------------------------------------------------------------


Private Sub Form_Load()
    Me.AutoResize = True
    Me.FitToScreen = True
    DoCmd.Maximize
    cmdResetForm_Click
    txtCaseID.SetFocus
    Set db = CurrentDb
End Sub
Private Sub Form_Open(Cancel As Integer)
Dim strRefCaseID As String
    If Not IsNull(Forms!frmHICSUpdate.OpenArgs) Then
        strRefCaseID = Forms!frmHICSUpdate.OpenArgs
        If Len(strRefCaseID) > 0 Then
            txtCaseID.Value = strRefCaseID
            cmdSearchHICSCase_Click
        End If
    End If
End Sub
Private Sub Form_Close()
    'clean up memory upon close
    If Not db Is Nothing Then
        Set db = Nothing
    End If
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
End Sub
Private Sub cmdOpenDisputeForm_Click()
    DoCmd.OpenForm "frmERRDispute", acNormal, , , , acDialog, txtCaseID.Value & "|" & Me.Form.Name
End Sub
'opens tblHICS_RCNO_Members, submits case data as new record
Private Sub cmdRCNODispute_Click()
    
    If CheckSubmission(2) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            txtErrorBar.SetFocus
            Exit Sub
        End If
    End If
    
On Error GoTo err1:
    Call CheckDB
    Dim strComment As String
    Set rs = db.OpenRecordset("tblHICS_RCNO_Members", dbOpenDynaset)
    If MsgBox("Do you want to refer this case to RECON team?", vbQuestion + vbYesNo, "Refer to RECON?") = vbYes Then
        strComment = InputBox("Please Enter a Comment for the Dispute", "Dispute Comment")
        If Len(strComment) = 0 Then
            MsgBox ("Action Cancelled")
        Else
            rs.AddNew
                rs![caseID] = txtCaseID.Value
                rs![memberId] = txtMemberID.Value
                rs![FirstName] = txtFirstName.Value
                rs![LastName] = txtLastName.Value
                rs![comment] = strComment
                rs![employeeId] = txtEmpID.Value
                rs![SubmitDate] = Now()
            rs.Update
            rs.Close
            MsgBox "Submitted to RECON team to work"
        End If
    Else
        Exit Sub
    End If
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmHICSUpdate; cmdRCNODispute_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
End Sub
Private Sub btnOpenSecondHICSForm_Click()
   DoCmd.OpenForm "frmHICSUpdate2", acNormal, , , , acWindowNormal
End Sub
Private Sub cmdViewPAF_Click()
    DoCmd.OpenForm "frmPAF", acNormal, , , , acWindowNormal, txtMemberID.Value
End Sub
Private Sub cmdViewPrevNotes_Click()
    If DCount("casework_notes", "tblCaseWorkNotes", "case_id = '" & txtCaseID.Value & "'") <> 0 Then
        DoCmd.OpenForm "frmHICSPriorNotes", acNormal, , , , acWindowNormal
    Else
        MsgBox ("No prior casework notes were found.")
    End If
End Sub
Private Sub cmdViewRCNO_Click()
    DoCmd.OpenForm "frmRCNOView", acNormal, , , , acWindowNormal, txtMemberID.Value
End Sub
Public Sub cmdResetForm_Click()
'CLEARS/RESETS ENTIRE FORM
    ClearForm
    ClearRes
    txtCaseID.Locked = False
    txtCaseID.SetFocus
    txtCaseStatus.BackColor = RGB(245, 246, 246)
    txtErrorBar.BackColor = RGB(245, 246, 246)
    cboResolveCase.Value = Null
    Deactivate cboCaseAction
    Deactivate txtDateReferred
    lblDateReferred.Visible = False
    Deactivate txtPermissionNote
    lblPermissionNote.Visible = False
    bolCaseResolved = False
    cboPendingHelpDesk.Value = "N"
    cboCallBack.Value = "NO"
    Call txtCaseworkNotes_Exit(0)
    Call txtInternalNotes_Exit(0)
    Call txtERRNoteReview_Change
    strInteraction = Null & vbNullString
    strCaseType = Null & vbNullString
    strCaseAction = Null & vbNullString
End Sub
Private Sub cmdClearNotes_Click()
'CLEARS NOTE SECTION
    If MsgBox("Do you want to clear all notes?", vbQuestion + vbYesNo, "Clear Notes") = vbYes Then
        Select Case bolCaseResolved
            Case True
                txtInternalNotes.Value = Null
                Call txtInternalNotes_Exit(0)
            Case False
                cboInteractionType = Null
                cboCaseType = Null
                cboResolveCase = Null
                cboCaseAction = Null
                cboCaseAction.Visible = False
                Deactivate txtERRNoteReview
                cmdOpenDisputeForm.Visible = False
                txtCaseworkNotes = Null
                txtInternalNotes = Null
                Deactivate txtResponseSentence
                Deactivate txtCMSResolutionStatement
                Deactivate txtResolutionOutcomeOtherText
                Deactivate txtDuplicateCaseID
                Deactivate cboTransferType
                Deactivate optAppealForm
                lblResponseSentence.Visible = False
                lblCMSResolutionStatement.Visible = False
                lblResolutionOutcomeOtherText.Visible = False
                lblDuplicateCaseID.Visible = False
                lblTransferType.Visible = False
                lblAppealForm.Visible = False
                Call txtCaseworkNotes_Exit(0)
                Call txtInternalNotes_Exit(0)
                Call txtERRNoteReview_Change
                strInteraction = Null & vbNullString
                strCaseType = Null & vbNullString
                strCaseAction = Null & vbNullString
        End Select
    Else
        Exit Sub
    End If
End Sub






'--------------------------------------------------------------SEARCH FUNCTIONS----------------------------------------------------------------------------------------------------------------------------------------
'ROSTER SEARCH
Private Sub cmdRosterSearch_Click()
    
    'VALIDATION
    If CheckSubmission(5) = True Then
        If MsgBox("Error: see 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            txtErrorBar.SetFocus
            Exit Sub
        End If
    End If
    
On Error GoTo err1:
    Call CheckDB
    Dim strSearch As String: strSearch = "[MemberNbr] = " & "'" & txtMemberID.Value & "'"
    Set rs = db.OpenRecordset("tblMemberRoster", dbOpenSnapshot)
    
    With rs
        .FindFirst (strSearch)
        If Not .NoMatch Then
            txtFirstName.Value = ![MembersFirstName]
            txtLastName.Value = ![MembersLastName]
            txtStreetAddress.Value = ![Address1]
            txtStreetAddress2.Value = ![Address2]
            txtCity.Value = ![city]
            txtState.Value = ![State]
            txtZIpCode.Value = ![ZipCode]
        Else: MsgBox ("Member not found on roster")
        End If
        .Close
    End With
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmHICSUpdate; cmdRosterSearch_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select

End Sub
'CASE SEARCH
Public Sub cmdSearchHICSCase_Click()
    
    If CheckSubmission(2) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            txtErrorBar.SetFocus
            Exit Sub
        End If
    End If

On Error GoTo err1:
    Call CheckDB
    Dim strSearch As String: strSearch = "[case_id] = '" & txtCaseID.Value & "'"
    Dim strMessage As String
    Dim strTemp As String: strTemp = ""
    Dim regexStr As String
    
    With Me
        '########################## CASE STATUS #################################
        Set rs = db.OpenRecordset("tblCaseStatus", dbOpenSnapshot)
        rs.FindFirst strSearch
        If rs.NoMatch Then
            strMessage = strMessage & "No Case Status Found." & vbCrLf
            rs.Close
        ElseIf rs![resolve_case] = "Y" Then
            MsgBox ("This case has been resolved. Only 'Internal Notes' can be submitted.")
            txtCaseStatus.Value = "Closed"
            txtCaseStatus.BackColor = RGB(255, 99, 71)
            txtERR.Value = rs![Err]
            rs.Close
            bolCaseResolved = True
            LockForm
        ElseIf rs![resolve_case] = "N" Then
            With rs
                cboResolveCase.Value = ![resolve_case]
                txtERR.Value = ![Err]
                cboPendingHelpDesk.Value = ![pending_helpdesk]
                If ![pending_helpdesk] = "Y" Then
                    Activate txtDateReferred
                    txtDateReferred.Value = ![date_referred]
                End If
                If ![updated_allowed] = True Then
                    cboCallBack = 1
                    Activate txtPermissionNote
                    txtPermissionNote.Value = ![perm_note]
                End If
                .Close
            End With
                txtCaseStatus.Value = "Open"
                txtCaseStatus.BackColor = RGB(144, 238, 144)
                txtCaseID.Locked = True
                cmdRosterSearch.Enabled = True
                bolCaseResolved = False
        End If
       ' ###################### HICS CASE DATA #######################################
        Set rs = db.OpenRecordset("tblCMSCaseData", dbOpenSnapshot)
        rs.FindFirst strSearch
        If rs.NoMatch Then
            strMessage = strMessage & "No Case Data found." & vbCrLf
        Else
            txtComplianceDueDate.Value = Format(rs![compliance_due_date], "m/d/yyyy")
            txtIssuerAssignDate.Value = Format(rs![issuer_assign_date], "m/d/yyyy")
            txtCaseLevel.Value = rs![case_level]
            txtLanguage.Value = rs![preferred_language]
            txtCongressionalCase.Value = rs![congressional_case]
            txtSubcategory.Value = rs![Subcategory]
                'call regex functions to reformat narrative
                If Not IsNull(rs![case_narrative]) And Not IsNull(rs![Subcategory]) Then
                    regexStr = RegexSwitch(rs![Subcategory])
                    txtCaseNarrative.Value = RegexNarrative(rs![case_narrative], regexStr)
                End If
            rs.Close
        End If
        '###################### MEMBER DEMOGRAPHICS #################################
        Set rs = db.OpenRecordset("tblMemberDemo", dbOpenSnapshot)
        rs.FindFirst strSearch
        If rs.NoMatch Then
            strMessage = strMessage & "New Case; No Member Information associated with case." & vbCrLf
        Else
            With rs
                txtMemberID.Value = ![member_id]
                txtFirstName.Value = ![first_name]
                txtLastName.Value = ![last_name]
                txtStreetAddress.Value = ![street_address]
                txtStreetAddress2.Value = ![street_address_2]
                txtCity.Value = ![city]
                txtState.Value = ![State]
                txtZIpCode.Value = ![zip_code]
                rs.Close
            End With
        End If
       
        '####################### COMMENTS FROM HICS SYSTEM ######################
        'case can have multiple comments - this concats them in a single string
        Set rs = db.OpenRecordset("tblComments", dbOpenSnapshot)
            rs.FindFirst strSearch
            Do While rs.EOF = False
                If Not rs.NoMatch Then
                    strTemp = "<b>Type: </b>" & rs![comment_type] & "<br>" & _
                    "<b>Date: </b>" & rs![created_date] & "<br>" & _
                    "<b>Comment: </b>" & Trim(rs![comment]) & "<br><br>" & strTemp
                ElseIf rs.NoMatch Then
                    rs.Close
                    Exit Do
                End If
                rs.FindNext strSearch
            Loop
        txtAddCaseNotes.Value = strTemp
        
        '###################### CONTACT INFO FROM HICS #############################
        Set rs = db.OpenRecordset("tblContact", dbOpenSnapshot)
        rs.FindFirst strSearch
        If rs.NoMatch Then
        strMessage = strMessage & "No CMS provided Contact Information." & vbCrLf
        Else
            With rs
                txtDemographic.Value = _
                    "<b>Name: </b>" & ![first_name] & " " & ![last_name] & "<br><br>" & _
                    "<b>Address: </b>" & ![Address] & " " & ![address_2] & "<br>" & _
                            ![city] & ", " & ![State] & " " & ![zip_code] & "<br><br>" & _
                    "<b>E-mail: </b>" & ![email] & "<br><br>" & _
                    "<b>DOB: </b>" & ![dob]
                txtPhoneNumber.Value = ![phone]
                txtAltPhoneNumber.Value = ![alt_phone]
                rs.Close
            End With
        End If
        
        '###################### ERR&R NOTES #############################
        Set rs = db.OpenRecordset("tblERRNotes", dbOpenSnapshot)
        rs.FindFirst ("CaseID = '" & txtCaseID.Value & "'")
        If rs.NoMatch Then
        Else
            With rs
                txtERRNotes.Value = _
                    "<b>Submit Date: </b>" & ![SubmitDate] & "<br><br>" & Replace(![Note], "ER&R", "<b>ER&R</b>")
                rs.Close
            End With
        End If
        
        '############################Set User ###############################
        strUser = fOSUserName()
        Set rs = db.OpenRecordset("Select * From tblEmployeeNames Where tblEmployeeNames.NetworkID = " & "'" & strUser & "';", dbOpenSnapshot)
        rs.MoveFirst
        strUser = rs![employee_id]
        txtEmpID.Value = strUser
        
        '############################ CLEAN UP ###############################
        If Not rs Is Nothing Then
            rs.Close
        End If
    End With
        
    If strMessage <> "" Then
        If MsgBox(strMessage, vbExclamation + vbOKOnly, "No Results") = vbOK Then: Exit Sub
    End If

err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmHICSUpdate; cmdSearchHICSCase_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
End Sub









'-------------------------------------------------------------- SUBMITS ----------------------------------------------------------------------------------------------------------------------------------------
Private Sub cmdSubmitForm_Click()

    '-- CHECK IF CLOSED --
    If bolCaseResolved = False And cboResolveCase.Value = -1 Then 'RESOLVE CASE
        SubmitResolutionNote
        'DoCmd.OpenForm "frmHICSConfirmSubmit", acNormal, , , , acDialog
        'Call SubmitResolutionNote
    ElseIf bolCaseResolved = False And cboResolveCase.Value = 0 Then 'NEW CASEWORK NOTE
        SubmitCaseWorkNote
        'DoCmd.OpenForm "frmHICSConfirmSubmit", acNormal, , , , acDialog
        'Call SubmitCaseWorkNote
    ElseIf bolCaseResolved = True Then
        SubmitInternalNote
    End If
    
End Sub
'NEW CASEWORK NOTE
Public Sub SubmitCaseWorkNote()

    'VALIDATION
    If CheckSubmission(1) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            txtErrorBar.SetFocus
            Exit Sub
        End If
    End If

    If MsgBox("Do you want to submit this casework note?", vbYesNo, "Submit Note?") = vbNo Then Exit Sub
    
    'level 1s get prompted if user agrees with case level. if disagree, add reason why
    If txtCaseLevel.Value = 1 And DLookup("issuer_level_agree", "tblCaseStatus", "case_id = '" & txtCaseID.Value & "'") = "Y" Then
        If MsgBox("Level 1 Case:" & vbCrLf & vbCrLf & "Do you agree with case level?", vbQuestion + vbYesNo, "Issuer Level Agree?") = vbNo Then
            strLevelDisagree = InputBox("Please Enter why you disagree", "Issuer Level Disagree")
        End If
    End If
    
'BEGIN WRITING DATA
On Error GoTo err1:
        Call CheckDB
        Dim rsPrevDemo As Recordset
        Dim crtDate As Date: crtDate = Now
        Dim caseWorkNoteId As Long
        Dim strSearch As String: strSearch = "[case_id] = " & "'" & txtCaseID.Value & "'"
        With Me
        
'################################## tblCaseStatus ##########################################
        Set rs = db.OpenRecordset("tblCaseStatus", dbOpenDynaset)
        rs.FindFirst (strSearch)
        If rs.NoMatch Then
            MsgBox ("Case not Found Contact Database Admin. ERROR IN CASE STATUS SUBMIT")
            Call LogError(0 & " " & "Case not found in tblCaseStatus", "frmHICSUpdate; SubmitCaseWorkNote()")
            Exit Sub
        ElseIf Not rs.NoMatch Then
            With rs
                If (![pending_helpdesk] <> cboPendingHelpDesk.Value) And _
                    (![pending_helpdesk] = "N") Then
                    .Edit
                    ![pending_helpdesk] = cboPendingHelpDesk.Value
                    ![date_referred] = txtDateReferred.Value
                    .Update
                End If
                
                If (txtCaseLevel.Value = 1 And strLevelDisagree <> "" And ![issuer_level_agree] = "Y") Then
                    .Edit
                    ![issuer_level_agree] = "N"
                    ![issuer_level_disagree_text] = strLevelDisagree
                    .Update
                End If
                
                If ![Err] = "N" And txtERR.Value = "Y" Then
                    .Edit
                    ![Err] = txtERR.Value
                    ![err_date] = Format(crtDate, "mm/dd/yyyy")
                    .Update
                ElseIf ![Err] = "Y" And txtERR.Value = "N" Then
                    .Edit
                    ![Err] = txtERR.Value
                    ![err_returned] = Date
                    .Update
                End If
                
                If cboCallBack.Value = 0 Then
                    .Edit
                    ![updated_allowed] = 0
                    .Update
                ElseIf cboCallBack.Value = 1 Then
                    .Edit
                    ![updated_allowed] = 1
                    ![perm_note] = txtPermissionNote.Value
                    ![perm_date] = crtDate
                    ![perm_specialist] = CInt(txtEmpID.Value)
                    .Update
                End If

            End With
            rs.Close
        End If
        
'########################### MEMBER DEMOGRAPHIC INFO ######################################
'matches on caseid, if case is already on tblMemberDemo

        Set rs = db.OpenRecordset("tblMemberDemo", dbOpenDynaset)
        rs.FindFirst (strSearch)
        If rs.NoMatch Then
            With rs
                .AddNew
                ![case_id] = txtCaseID.Value
                ![member_id] = txtMemberID.Value
                ![first_name] = txtFirstName.Value
                ![last_name] = txtLastName.Value
                ![street_address] = txtStreetAddress.Value
                ![street_address_2] = txtStreetAddress2.Value
                ![city] = txtCity.Value
                ![State] = txtState.Value
                ![zip_code] = txtZIpCode.Value
                ![employee_id] = CInt(txtEmpID.Value)
                .Update
                .Close
            End With
        ElseIf Not rs.NoMatch Then
            With rs
                If (![member_id] <> txtMemberID.Value) Or _
                    (![first_name] <> txtFirstName.Value) Or _
                    (![last_name] <> txtLastName.Value) Or _
                    (![street_address] <> txtStreetAddress.Value) Or _
                    (![street_address_2] <> txtStreetAddress2.Value) Or _
                    (![city] <> txtCity.Value) Or _
                    (![State] <> txtState.Value) Or _
                    (![zip_code] <> txtZIpCode.Value) Then
                   Set rsPrevDemo = db.OpenRecordset("tblPreviousMemberDemo", dbOpenDynaset)
                        rsPrevDemo.AddNew
                        rsPrevDemo![case_id] = ![case_id]
                        rsPrevDemo![member_id] = ![member_id]
                        rsPrevDemo![first_name] = ![first_name]
                        rsPrevDemo![last_name] = ![last_name]
                        rsPrevDemo![street_address] = ![street_address]
                        rsPrevDemo![street_address_2] = ![street_address_2]
                        rsPrevDemo![city] = ![city]
                        rsPrevDemo![State] = ![State]
                        rsPrevDemo![zip_code] = ![zip_code]
                        rsPrevDemo![date_changed] = crtDate
                        rsPrevDemo![employee_id] = CInt(txtEmpID.Value)
                        rsPrevDemo![prev_employee_id] = ![employee_id]
                        rsPrevDemo.Update
                        rsPrevDemo.Close
                        Set rsPrevDemo = Nothing
                        
                        .Edit
                        ![member_id] = txtMemberID.Value
                        ![first_name] = txtFirstName.Value
                        ![last_name] = txtLastName.Value
                        ![street_address] = txtStreetAddress.Value
                        ![street_address_2] = txtStreetAddress2.Value
                        ![city] = txtCity.Value
                        ![State] = txtState.Value
                        ![zip_code] = txtZIpCode.Value
                        ![employee_id] = CInt(strUser)
                        .Update
                    End If
                .Close
            End With
        End If
        
    '######################################CASEWORK NOTES ################################################
    'always adds new record
        Set rs = db.OpenRecordset("tblCaseWorkNotes", dbOpenDynaset)
            With rs
                .AddNew
                ![case_id] = txtCaseID.Value
                ![casework_notes] = Clean(txtCaseworkNotes.Value)
                ![submit_date] = crtDate
                ![is_res] = "N"
                ![employee_id] = CInt(txtEmpID.Value)
                ![interaction_type] = cboInteractionType.Value
                ![case_action] = cboCaseAction.Value
                ![congressional_case] = txtCongressionalCase.Value
                .Update
                
                'GET ID OF NEWLY SUBMITTED RECORD
                .Move 0, .LastModified
                caseWorkNoteId = ![Id]
                
                .Close
          End With
     
     '####################################Internal System Notes ###########################################
     'adds new record
        If Nz(txtInternalNotes.Value, "") <> "" Then
            Set rs = db.OpenRecordset("tblInternalNotes", dbOpenDynaset)
                With rs
                    .AddNew
                    ![case_id] = txtCaseID.Value
                    ![Note] = txtInternalNotes.Value
                    ![submit_date] = crtDate
                    ![employee_id] = CInt(txtEmpID.Value)
                    ![casework_note_id] = caseWorkNoteId
                    .Update
                    .Close
                End With
        End If
        
    '#################################### ERR Note ###########################################
    'open tblERRNotes table, append new record. all are direct disputes, sent to hics system
    If cboCaseAction.Value = 7 And Not IsNull(txtERRNoteReview) Then
        Set rs = db.OpenRecordset("tblERRNotes", dbOpenDynaset)
            With rs
                rs.AddNew
                rs![caseID] = txtCaseID.Value
                rs![Note] = txtERRNoteReview.Value
                rs![SubmitDate] = crtDate
                rs![employeeId] = CInt(strUser)
                rs![casework_note_id] = caseWorkNoteId
                rs.Update
                rs.Close
            End With
    End If
    
    '#####################################CUT LOG NOTES ##################################################
            'only add 1 cutlog for open case to indicate taht casework is in progress
            Set rs = db.OpenRecordset("tblCutlogNotes", dbOpenDynaset)
            rs.FindFirst (strSearch)
            If rs.NoMatch Then
                With rs
                    .AddNew
                        ![case_id] = txtCaseID.Value
                        ![activity_code] = "HCS010"
                        ![cutlog_notes] = "HICS Case " & txtCaseID.Value & " has been received and is currently in progress. " & _
                            "Please contact Marketplace Enrollment for info."
                        ![submit_date] = crtDate
                        ![uploaded] = "N"
                        ![employee_id] = CInt(txtEmpID.Value)
                        ![casework_note_id] = caseWorkNoteId
                    .Update
                    .Close
                End With
            End If
                    
    End With
    
    MsgBox ("Note submitted")
    cmdResetForm_Click
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmHICSUpdate; SubmitCaseWorkNote()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub
Public Sub SubmitResolutionNote()

'VALIDATION
    If CheckSubmission(4) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            txtErrorBar.SetFocus
            Exit Sub
        End If
    End If
    
'BEGIN WRITING DATA
On Error GoTo err1:
    If MsgBox("Do you wish to resolve the case?", vbYesNo, "Resolve Case?") = vbNo Then Exit Sub
    Call CheckDB
    Dim caseWorkNoteId As Long
    Dim crtDate As Date: crtDate = Now()
    Dim strSearch As String: strSearch = "[case_id] = " & "'" & txtCaseID.Value & "'"
    
    '########################### MEMBER DEMOGRAPHIC INFO ######################################
    'matches on caseid, if case is already on tblMemberDemo
    Set rs = db.OpenRecordset("tblMemberDemo", dbOpenDynaset)
    rs.FindFirst (strSearch)
    If rs.NoMatch Then
        With rs
            .AddNew
            ![case_id] = txtCaseID.Value
            ![member_id] = txtMemberID.Value
            ![first_name] = txtFirstName.Value
            ![last_name] = txtLastName.Value
            ![street_address] = txtStreetAddress.Value
            ![street_address_2] = txtStreetAddress2.Value
            ![city] = txtCity.Value
            ![State] = txtState.Value
            ![zip_code] = txtZIpCode.Value
            ![employee_id] = CInt(txtEmpID.Value)
            .Update
            .Close
        End With
    ElseIf Not rs.NoMatch Then
        With rs
            If (![member_id] <> txtMemberID.Value) Or _
                (![first_name] <> txtFirstName.Value) Or _
                (![last_name] <> txtLastName.Value) Or _
                (![street_address] <> txtStreetAddress.Value) Or _
                (![street_address_2] <> txtStreetAddress2.Value) Or _
                (![city] <> txtCity.Value) Or _
                (![State] <> txtState.Value) Or _
                (![zip_code] <> txtZIpCode.Value) Then
               Set rsPrevDemo = db.OpenRecordset("tblPreviousMemberDemo", dbOpenDynaset)
                    rsPrevDemo.AddNew
                    rsPrevDemo![case_id] = ![case_id]
                    rsPrevDemo![member_id] = ![member_id]
                    rsPrevDemo![first_name] = ![first_name]
                    rsPrevDemo![last_name] = ![last_name]
                    rsPrevDemo![street_address] = ![street_address]
                    rsPrevDemo![street_address_2] = ![street_address_2]
                    rsPrevDemo![city] = ![city]
                    rsPrevDemo![State] = ![State]
                    rsPrevDemo![zip_code] = ![zip_code]
                    rsPrevDemo![date_changed] = crtDate
                    rsPrevDemo![employee_id] = CInt(txtEmpID.Value)
                    rsPrevDemo![prev_employee_id] = ![employee_id]
                    rsPrevDemo.Update
                    rsPrevDemo.Close
                    Set rsPrevDemo = Nothing
                    
                    .Edit
                    ![member_id] = txtMemberID.Value
                    ![first_name] = txtFirstName.Value
                    ![last_name] = txtLastName.Value
                    ![street_address] = txtStreetAddress.Value
                    ![street_address_2] = txtStreetAddress2.Value
                    ![city] = txtCity.Value
                    ![State] = txtState.Value
                    ![zip_code] = txtZIpCode.Value
                    ![employee_id] = CInt(strUser)
                    .Update
                End If
            .Close
        End With
    End If
    
    '#################################### CASEWORK NOTES ###########################################
    Set rs = db.OpenRecordset("tblCaseWorkNotes", dbOpenDynaset)
        rs.FindFirst "[case_id] = " & "'" & txtCaseID.Value & "' AND [is_res] = 'Y'"

        'check if there are existing resolutions on tbl. If so, set [prev_res] = -1 before adding new record
        If Not rs.NoMatch Then
            With rs
                Do While Not .NoMatch
                    If ![prev_res] = 0 Then
                        .Edit
                        ![prev_res] = -1
                        .Update
                    End If
                    .FindNext "[case_id] = " & "'" & txtCaseID.Value & "' AND [is_res] = 'Y'"
                Loop
            End With
        End If
        
        'add new record
        With rs
            .AddNew
                ![case_id] = txtCaseID.Value
                ![casework_notes] = Clean(txtCaseworkNotes.Value)
                ![submit_date] = crtDate
                ![is_res] = "Y"
                ![resolution_response] = txtResponseSentence.Value
                ![resolution_date] = Format(Date, "mm/dd/yyyy")
                ![resolution_outcome] = txtResolutionOutcome.Value
                ![resolution_outcome_other_text] = Nz(txtResolutionOutcomeOtherText.Value, "")
                ![employee_id] = CInt(txtEmpID.Value)
                ![interaction_type] = cboInteractionType.Value
                ![case_action] = cboCaseAction.Value
                ![congressional_case] = txtCongressionalCase.Value
                ![appeal_form] = optAppealForm.Value
            .Update
            
            'GET ID OF NEWLY SUBMITTED RECORD
            .Move 0, .LastModified
            caseWorkNoteId = ![Id]
            
            .Close
        End With
    
'#################################### CASE STATUS ###########################################
        Set rs = db.OpenRecordset("tblCaseStatus", dbOpenDynaset)
        rs.FindFirst (strSearch)
        If rs.NoMatch Then
            MsgBox ("Case not Found Contact Database Admin. ERROR IN CASE STATUS SUBMIT")
            Call LogError(0 & " " & "Case not found in tblCaseStatus", "frmHICSUpdate; SubmitResolutionNote()")
            Exit Sub
        Else
            With rs
                .Edit
                    ![resolve_case] = "Y"
                    
                    If cboCaseAction.Value = 5 Then 'DUPLICATE
                        ![duplicate] = -1
                        ![duplicate_caseID] = txtDuplicateCaseID.Value
                    End If
                    
                    If cboCaseAction.Value = 9 Then 'TRANSFERRED
                        ![transferred] = -1
                        ![trasnfered_to] = cboTransferType.Value
                    End If
                    
                    If ![Err] = "Y" Then
                        ![Err] = "N"
                        ![err_returned] = Format(crtDate, "mm/dd/yyyy")
                    End If
                .Update
                rs.Close
            End With
        End If
        
'##################################### CUT LOG NOTES ##################################################
        'cutlog required to indicate closing of case
        Set rs = db.OpenRecordset("tblCutlogNotes", dbOpenDynaset)
            With rs
                .AddNew
                    ![case_id] = txtCaseID.Value
                    ![activity_code] = "HCS010"
                    If cboCaseAction.Value <> 9 Then 'for non-transferred cases
                        ![cutlog_notes] = Left(txtResponseSentence, 187)
                    ElseIf cboCaseAction.Value = 9 Then
                        ![cutlog_notes] = "Case transferred to " & cboTransferType
                    End If
                    ![submit_date] = crtDate
                    ![uploaded] = "N"
                    ![employee_id] = CInt(txtEmpID.Value)
                    ![casework_note_id] = caseWorkNoteId
                .Update
                .Close
            End With
            
'####################################Internal System Notes ###########################################
     'adds new record
        If Nz(txtInternalNotes.Value, "") <> "" Then
            Set rs = db.OpenRecordset("tblInternalNotes", dbOpenDynaset)
                With rs
                    .AddNew
                    ![case_id] = txtCaseID.Value
                    ![Note] = txtInternalNotes.Value
                    ![submit_date] = crtDate
                    ![employee_id] = CInt(txtEmpID.Value)
                    ![casework_note_id] = caseWorkNoteId
                    .Update
                    .Close
                End With
        End If
            
'##################################### TRENDING TABLE ##################################################
        Set rs = db.OpenRecordset("tblResolutionOutcome_Category", dbOpenDynaset)
            rs.AddNew
                rs![case_id] = txtCaseID.Value
                rs![submit_date] = crtDate
                rs![category] = cboCaseType.Column(1)
                If cboCaseAction.Value <> 9 Then 'for non-transferred cases
                    rs![outcome] = txtOutcome.Value
                ElseIf cboCaseAction.Value = 9 Then
                    rs![outcome] = "Transferred"
                End If
                rs![employee_id] = CInt(txtEmpID.Value)
            rs.Update
            rs.Close

        cmdResetForm_Click
        MsgBox "Resolution Summary Submitted"

err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmHICSUpdate; SubmitResolutionNote()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub
Public Sub SubmitInternalNote()
    If CheckSubmission(1) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            txtErrorBar.SetFocus
            Exit Sub
        End If
    End If
    
'BEGIN WRITING DATA
On Error GoTo err1:
    If MsgBox("Submit internal note on closed case?", vbYesNo, "Internal Note") = vbNo Then Exit Sub
    Call CheckDB
    Dim caseWorkNoteId As Long
    Dim rs As DAO.Recordset
    
    '####################################Internal System Notes ###########################################
    Set rs = db.OpenRecordset("SELECT MAX(ID) FROM tblCaseWorkNotes WHERE case_id = '" & txtCaseID.Value & "' And is_res = 'Y';")
    With rs
        If .recordCount > 0 Then
            caseWorkNoteId = .Fields(0)
        End If
    End With
    rs.Close
    
    Set rs = db.OpenRecordset("tblInternalNotes", dbOpenDynaset)
    With rs
        .AddNew
            ![case_id] = txtCaseID.Value
            ![Note] = txtInternalNotes.Value
            ![submit_date] = Now()
            ![employee_id] = CInt(txtEmpID.Value)
            ![casework_note_id] = IIf(Not IsNull(caseWorkNoteId), caseWorkNoteId, Null)
        .Update
    End With
    rs.Close
    
    cmdResetForm_Click
    MsgBox ("Note submitted")
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmHICSUpdate; SubmitInternalNote()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
End Sub








'--------------------------------------------------------------CONTROL ACTIONS----------------------------------------------------------------------------------------------------------------------------------------

Private Sub cboCallBack_AfterUpdate()
 If cboCallBack.Value = 1 Then
        Activate txtPermissionNote
        lblPermissionNote.Visible = True
    Else: Deactivate txtPermissionNote: lblPermissionNote.Visible = False
    End If
End Sub
Private Sub cboPendingHelpDesk_Change()
    If Not IsNull(cboPendingHelpDesk.Value) Then
        Select Case cboPendingHelpDesk.Value
            Case "Y"
                Activate txtDateReferred
                lblDateReferred.Visible = True
                txtDateReferred.Value = Date
            Case "N"
                Deactivate txtDateReferred
            Case Else:
        End Select
    Else
    End If
End Sub
Private Sub txtCaseworkNotes_KeyDown(KeyCode As Integer, Shift As Integer)
    txtCaseworkNotes.ForeColor = vbBlack
End Sub
Private Sub txtCaseworkNotes_Exit(Cancel As Integer)
    If IsNull(txtCaseworkNotes) Or txtCaseworkNotes.Value = "" Then
        txtCaseworkNotes.ForeColor = RGB(166, 166, 166)
    End If
End Sub
Public Sub txtERRNoteReview_Change()
    If IsNull(txtERRNoteReview) Or txtERRNoteReview.Value = "" Then
        txtERRNoteReview.ForeColor = RGB(166, 166, 166)
    Else
        txtERRNoteReview.ForeColor = vbBlack
    End If
End Sub
Private Sub txtInternalNotes_Exit(Cancel As Integer)
    If IsNull(txtInternalNotes) Or txtInternalNotes.Value = "" Then
        txtInternalNotes.ForeColor = RGB(166, 166, 166)
    End If
End Sub
Private Sub txtInternalNotes_KeyDown(KeyCode As Integer, Shift As Integer)
    txtInternalNotes.ForeColor = vbBlack
End Sub
Private Sub txtDuplicateCaseID_Exit(Cancel As Integer)
    Dim dupeID As String: dupeID = txtDuplicateCaseID.Value & vbNullString
    txtResponseSentence.Value = "Case was resolved as a duplicate of case " & dupeID
    txtResolutionOutcomeOtherText = "Case was resolved as a duplicate of case " & dupeID
End Sub
Private Sub cboInteractionType_Change()
    Call SplitNotes(1)
    Call txtCaseworkNotes_KeyDown(0, 0)
End Sub
Private Sub cboCaseType_Change()
    Call SplitNotes(2)
    Call txtCaseworkNotes_KeyDown(0, 0)
    Call GetCaseActList
End Sub
Private Sub cboResolveCase_AfterUpdate()
    Dim statement As String
    If Not IsNull(cboResolveCase) Then
        Activate cboCaseAction
        cboCaseAction.Value = Null
        lblCaseAction.Visible = True
        ClearRes
        Call SplitNotes(3)
        Call GetCaseActList
    Else
        Deactivate cboCaseAction
        cboResolveCase.Value = 0
        MsgBox ("Value must be 'Y' or 'N'.")
    End If
End Sub
Private Sub cboCaseAction_Change()
    ClearRes
    Select Case cboCaseAction.Value
        Case 1, 2, 3, 4, 10, 11, 12, 14, 15 'standard resolutions, coe, ue
            Activate txtResponseSentence
            Activate txtCMSResolutionStatement
            Activate optAppealForm
            lblResponseSentence.Visible = True
            lblCMSResolutionStatement.Visible = True
            lblAppealForm.Visible = True
        Case 5 'duplicate
            Activate txtCMSResolutionStatement
            Activate txtResolutionOutcomeOtherText
            Activate txtDuplicateCaseID
            lblCMSResolutionStatement.Visible = True
            lblResolutionOutcomeOtherText.Visible = True
            lblDuplicateCaseID.Visible = True
        Case 6, 8 'non-resolution

        Case 7 'ER&R
            Activate txtERRNoteReview
            cmdOpenDisputeForm.Visible = True
        Case 9 'transferred
            txtCaseworkNotes.Value = Null
            Activate cboTransferType
            lblTransferType.Visible = True
    End Select
    Call SplitNotes(3)
    Call txtCaseworkNotes_KeyDown(0, 0)
End Sub

'REFERRENCES NEEDED FOR CALLS TO WINDOWS SHELL MESSAGE BOX:
    'Windows Script Host Object Model
    'WSHControllerLibrary
    
    
Private Sub txtAddCaseNotes_DblClick(Cancel As Integer)
    Dim commStr As String
    commStr = txtAddCaseNotes.Value & vbNullString
    If Not IsNull(commStr) And commStr <> "" Then
        commStr = Replace(commStr, "<b>", "")
        commStr = Replace(commStr, "</b>", "")
        commStr = Replace(commStr, "<br>", vbCrLf)
        Dim msg As String: msg = "Additional Case Comments"
        Dim WSHShell
        Set WSHShell = CreateObject("WScript.Shell")
        WSHShell.PopUp commStr, 0, msg, 0
    End If
End Sub
Private Sub txtCaseNarrative_DblClick(Cancel As Integer)
    Dim narrStr As String
    narrStr = txtCaseNarrative.Value & vbNullString
    If Not IsNull(narrStr) And narrStr <> "" Then
        narrStr = Replace(narrStr, "<b>", "")
        narrStr = Replace(narrStr, "</b>", "")
        narrStr = Replace(narrStr, "<br>", vbCrLf)
        Dim msg As String: msg = "CMS Case Narrative"
        Dim WSHShell
        Set WSHShell = CreateObject("WScript.Shell")
        WSHShell.PopUp narrStr, 0, msg, 0
    End If
End Sub
Private Sub txtDemographic_DblClick(Cancel As Integer)
    Dim demoStr As String
    demoStr = txtDemographic.Value & vbNullString
    If Not IsNull(demoStr) And demoStr <> "" Then
        demoStr = Replace(demoStr, "<b>", "")
        demoStr = Replace(demoStr, "</b>", "")
        demoStr = Replace(demoStr, "<br>", vbCrLf)
        Dim msg As String: msg = "CMS Contact Info"
        Dim WSHShell
        Set WSHShell = CreateObject("WScript.Shell")
        WSHShell.PopUp demoStr, 0, msg, 0
    End If
End Sub








'--------------------------------------------------------------FUNCTIONS----------------------------------------------------------------------------------------------------------------------------------------

Public Function CheckSubmission(submitType As Integer) As Boolean

    Dim msgStr As String, headerStr As String, footerStr As String, resultStr As String
    Dim regex As RegExp: Set regex = New RegExp
    headerStr = "<ul>"
    footerStr = "</ul>"
    
    Select Case submitType
        Case 1 '--SUBMIT BUTTON--
            'UNRESOLVED
            If bolCaseResolved = False Then
                With regex
                    .MultiLine = False
                    .Global = True
                    .IgnoreCase = True
                    .Pattern = "(\[[^\[\]]+\])|\[|\]"
                End With
                If Not IsNull(txtCaseworkNotes) Then
                    If regex.test(txtCaseworkNotes.Value) = True Then
                        msgStr = msgStr & "<li><b>Casework Notes</b> contains placeholder text. Please update any information contained in [brackets].</li>"
                    End If
                End If
                If DLookup("resolve_case", "tblCaseStatus", "case_id = '" & txtCaseID.Value & "'") = "Y" Then
                    msgStr = msgStr & "<li>Cannot submit casework note on closed case.</li>"
                End If
                If txtCaseStatus = "Closed" Then
                    msgStr = msgStr & "<li><b>Submit Error.</b> Please close the form and reopen</li>"
                End If
                If txtCaseID.Value = "" Or IsNull(txtCaseID.Value) Then
                    msgStr = msgStr & "<li><b>Case ID</b> cannot be blank.</li>"
                End If
                If txtFirstName.Value = "" Or IsNull(txtFirstName.Value) = True Then
                    msgStr = msgStr & "<li><b>First Name</b> cannot be blank.</li>"
                End If
                If txtLastName.Value = "" Or IsNull(txtLastName.Value) = True Then
                    msgStr = msgStr & "<li><b>Last Name</b> cannot be blank.</li>"
                End If
                If txtStreetAddress.Value = "" Or IsNull(txtStreetAddress.Value) = True Then
                    msgStr = msgStr & "<li><b>Street 1</b> cannot be blank.</li>"
                End If
                If txtCity.Value = "" Or IsNull(txtCity.Value) = True Then
                    msgStr = msgStr & "<li><b>City</b> cannot be blank.</li>"
                End If
                If txtState.Value = "" Or IsNull(txtState.Value) = True Then
                    msgStr = msgStr & "<li><b>State</b> cannot be blank.</li>"
                End If
                If txtZIpCode.Value = "" Or IsNull(txtZIpCode.Value) = True Then
                    msgStr = msgStr & "<li><b>Zip Code</b> cannot be blank.</li>"
                End If
                If txtCaseworkNotes.Value = "" Or IsNull(txtCaseworkNotes.Value) = True Then
                    msgStr = msgStr & "<li><b>Casework Notes</b> cannot be blank.</li>"
                End If
                If Len(txtCaseworkNotes + txtERRNoteReview) > 4000 Then
                    msgStr = msgStr & "<li><b>Casework Notes</b> cannot exceed 4000 characters. Please edit and resubmit.</li>"
                End If
                If IsNull(cboResolveCase) Or cboResolveCase = "" Then
                    msgStr = msgStr & "<li><b>Resolve Case</b> cannot be blank.</li>"
                End If
                If (cboCallBack.Value = 1 And (txtPermissionNote.Value = "" Or IsNull(txtPermissionNote))) Then
                    msgStr = msgStr & "<li><b>Permission Note</b> cannot be blank.</li>"
                End If
                If cboCaseAction.Value = 7 And IsNull(txtERRNoteReview.Value) = True Then
                    msgStr = msgStr & "<li>You must create a dispute when <b>Case Action</b> is set to 'ER&R'.</li>"
                End If
                Dim errVal As String: errVal = DLookup("err", "tblCaseStatus", "case_id = '" & txtCaseID.Value & "'")
                If cboCaseAction.Value = 7 And errVal = "Y" Then
                    msgStr = msgStr & "<li>Case is already referred to 'ER&R'.</li>"
                End If
                If cboInteractionType.Value = "" Or IsNull(cboInteractionType) Then
                    msgStr = msgStr & "<li><b>Interaction Type</b> cannot be blank.</li>"
                End If
                If cboCaseType.Value = "" Or IsNull(cboCaseType) Then
                    msgStr = msgStr & "<li><b>Case Type</b> cannot be blank.</li>"
                End If
                If cboResolveCase.Value = "" Or IsNull(cboResolveCase) Then
                    msgStr = msgStr & "<li><b>Resolve Case</b> cannot be blank.</li>"
                End If
                If cboCaseAction.Value = "" Or IsNull(cboCaseAction) Then
                    msgStr = msgStr & "<li><b>Case Action</b> cannot be blank.</li>"
                End If
            'RESOLVED
            ElseIf bolCaseResolved = True Then
                If txtCaseStatus = "Open" Then
                    msgStr = msgStr & "<li><b>Submit Error.</b> Please close the form and reopen</li>"
                End If
                If IsNull(txtInternalNotes) Or txtInternalNotes.Value = "" Then
                    msgStr = msgStr & "<li><b>Internal Notes</b> cannot be blank when submitting a closed case.</li>"
                End If
            End If
            
        Case 2 '--SEARCH CASE -- SEARCH PREV NOTES -- REFER TO RECON --
            With regex
                .MultiLine = False
                .Global = True
                .IgnoreCase = True
                .Pattern = "^[a-zA-Z0-9]{11}$"
            End With
            If txtCaseID.Value = "" Or IsNull(txtCaseID.Value) = True Then
                msgStr = msgStr & "<li><b>Case ID</b> cannot be blank.</li>"
            End If
            If Not IsNull(txtCaseID) Then
                If regex.test(txtCaseID.Value) = False Then
                    msgStr = msgStr & "<li><b>Case ID</b> is invalid. Please try again.</li>"
                End If
            End If
        Case 3 '--EDIT PREV NOTES--
        Case 4 '--RESOLVE CASE--
            With regex
                .MultiLine = False
                .Global = True
                .IgnoreCase = True
                .Pattern = "(\[[^\[\]]+\])|\[|\]"
            End With
            If Not IsNull(txtCaseworkNotes) Then
                If regex.test(txtCaseworkNotes.Value) = True Then
                    msgStr = msgStr & "<li><b>Casework Notes</b> contains placeholder text. Please update any information contained in [brackets].</li>"
                End If
            End If
            If DLookup("resolve_case", "tblCaseStatus", "case_id = '" & txtCaseID.Value & "'") = "Y" Then
                msgStr = msgStr & "<li>Cannot submit casework note on closed case.</li>"
            End If
            If txtCaseID.Value = "" Or IsNull(txtCaseID.Value) Then
                msgStr = msgStr & "<li><b>Case ID</b> cannot be blank.</li>"
            End If
            If txtFirstName.Value = "" Or IsNull(txtFirstName.Value) = True Then
                msgStr = msgStr & "<li><b>First Name</b> cannot be blank.</li>"
            End If
            If txtLastName.Value = "" Or IsNull(txtLastName.Value) = True Then
                msgStr = msgStr & "<li><b>Last Name</b> cannot be blank.</li>"
            End If
            If txtStreetAddress.Value = "" Or IsNull(txtStreetAddress.Value) = True Then
                msgStr = msgStr & "<li><b>Street 1</b> cannot be blank.</li>"
            End If
            If txtCity.Value = "" Or IsNull(txtCity.Value) = True Then
                msgStr = msgStr & "<li><b>City</b> cannot be blank.</li>"
            End If
            If txtState.Value = "" Or IsNull(txtState.Value) = True Then
                msgStr = msgStr & "<li><b>State</b> cannot be blank.</li>"
            End If
            If txtZIpCode.Value = "" Or IsNull(txtZIpCode.Value) = True Then
                msgStr = msgStr & "<li><b>Zip Code</b> cannot be blank.</li>"
            End If
            If txtCaseworkNotes.Value = "" Or IsNull(txtCaseworkNotes.Value) = True Then
                msgStr = msgStr & "<li><b>Casework Notes</b> cannot be blank.</li>"
            End If
            If Len(txtCaseworkNotes + txtERRNoteReview) > 4000 Then
                msgStr = msgStr & "<li><b>Casework Notes</b> cannot exceed 4000 characters. Please edit and resubmit.</li>"
            End If
            If cboInteractionType.Value = "" Or IsNull(cboInteractionType) Then
                msgStr = msgStr & "<li><b>Interaction Type</b> cannot be blank.</li>"
            End If
            If cboCaseType.Value = "" Or IsNull(cboCaseType) Then
                msgStr = msgStr & "<li><b>Case Type</b> cannot be blank.</li>"
            End If
            If cboResolveCase.Value = "" Or IsNull(cboResolveCase) Then
                msgStr = msgStr & "<li><b>Resolve Case</b> cannot be blank.</li>"
            End If
            If cboCaseAction.Value = "" Or IsNull(cboCaseAction) Then
                msgStr = msgStr & "<li><b>Case Action</b> cannot be blank.</li>"
            End If
            If cboCaseAction.Value = 5 And (txtDuplicateCaseID.Value = "" Or IsNull(txtDuplicateCaseID)) Then
                msgStr = msgStr & "<li><b>Duplicate Case ID</b> cannot be blank.</li>"
            End If
            If cboCaseAction.Value = 9 And (cboTransferType.Value = "" Or IsNull(cboTransferType)) Then
                msgStr = msgStr & "<li><b>Transferred to Whom</b> cannot be blank.</li>"
            End If
            If IsNull(optAppealForm) And (cboCaseAction.Value = 1 Or cboCaseAction.Value = 2 Or cboCaseAction.Value = 3 _
            Or cboCaseAction.Value = 4 Or cboCaseAction = 10 Or cboCaseAction = 11 Or cboCaseAction = 12) Then
                msgStr = msgStr & "<li><b>Include Appeal Form</b> cannot be blank.</li>"
            End If
            If (cboCaseAction.Value = 1 Or cboCaseAction.Value = 2 Or cboCaseAction.Value = 3 _
            Or cboCaseAction.Value = 4) And IsNull(txtResponseSentence) Then
                msgStr = msgStr & "<li><b>Letter Response</b> cannot be blank.</li>"
            End If
            
            If (cboCaseAction.Value = 1 Or cboCaseAction.Value = 2 Or cboCaseAction.Value = 3 _
            Or cboCaseAction.Value = 4 Or cboCaseAction.Value = 14 Or cboCaseAction.Value = 15) And Not IsNull(txtResponseSentence) Then
                If regex.test(txtResponseSentence.Value) = True Then
                    msgStr = msgStr & "<li><b>Letter Response</b> contains placeholder text. Please update any information contained in [brackets].</li>"
                End If
                regex.Pattern = "([0-9]{1,2}\.[0-9]{1,2}\.[0-9]{2,4})"
                If regex.test(txtResponseSentence.Value) = True Then
                    msgStr = msgStr & "<li>Dates within the <b>Letter Response</b> must be in <b>'mm/dd/yyyy'</b> format.</li>"
                End If
            End If

            If (cboCaseAction.Value = 1 Or cboCaseAction.Value = 2 Or cboCaseAction.Value = 3 _
            Or cboCaseAction.Value = 4) And (Not IsNull(txtResponseSentence) And InStr(txtResponseSentence, "EXXXXXXXXXX")) Then
                msgStr = msgStr & "<li><b>Letter Response</b> contains placeholder case id (EXXXXXXXXXX). Please update with correct case id.</li>"
            End If
            
            If (InStr(txtCaseworkNotes.Value, "misrepresentation or misinformation") <> 0 Or InStr(txtCaseworkNotes.Value, "SEP*") <> 0) _
            And cboResolveCase.Value = "Y" Then
                 msgStr = msgStr & "<li>You cannot close a case involving misrepresentation for non-response; please see leadership.</li>"
            End If
            If (txtERR.Value = "Y" And cboResolveCase.Value = "Y") Then
                msgStr = msgStr & "<li>You cannot resolve a case that is referred to ERR.</li>"
            End If
        Case 5 '--SEARCH ROSTER--
            With regex
                .MultiLine = False
                .Global = True
                .IgnoreCase = True
                .Pattern = "^[0-9]{11}$"
            End With
            txtMemberID.Value = ScrubMemberId(txtMemberID.Value & vbNullString)
            If txtMemberID.Value = "" Or IsNull(txtMemberID.Value) = True Then
                msgStr = msgStr & "<li><b>Member ID</b> cannot be blank.</li>"
            End If
            If Not IsNull(txtMemberID) Then
                If regex.test(txtMemberID.Value) = False Then
                    msgStr = msgStr & "<li><b>Member ID</b> is invalid. Please try again.</li>"
                End If
            End If

    End Select

    If msgStr = "" Then 'no errors
        txtErrorBox.Value = Null
        txtErrorBar.Value = Null
        txtErrorBar.BackColor = RGB(245, 246, 246)
        CheckSubmission = False
    Else 'errors, show msg to user
        txtErrorBox.Value = headerStr & msgStr & footerStr
        txtErrorBar.Value = "Submission Errors"
        txtErrorBar.BackColor = RGB(255, 186, 0)
        CheckSubmission = True
    End If

End Function
Public Sub SplitNotes(position As Integer)
    '--GLOBAL VARS--
    'strInteraction, strCaseType, strCaseAction
    Dim statement As String, noteStr As String
    Dim counter As Integer
    Dim textArray As Variant
    
    'check that txtCaseworkNotes has 2 pipes
    noteStr = txtCaseworkNotes.Value & vbNullString
    For i = 0 To Len(noteStr) - 1
        If Mid(noteStr, i + 1, 1) = "|" Then
            counter = counter + 1
        End If
    Next i
    
    'split string on pipes, scrub break chars and assign to global vars
    If counter <> 2 Then
    Else
        textArray = Split(noteStr, "|")
        textArray(0) = Replace(textArray(0), vbCrLf, "")
        textArray(1) = Replace(textArray(1), vbCrLf, "")
        textArray(2) = Replace(textArray(2), vbCrLf, "")
        strInteraction = textArray(0)
        strCaseType = textArray(1)
        strCaseAction = textArray(2)
    End If
    
    'add new statement to global vars
    Select Case position
        Case 1
            strInteraction = DLookup("statement", "tblHICSInteractionType", "ID = " & cboInteractionType.Value)
            If cboCaseAction.Value = 5 Or cboCaseAction.Value = 9 Then: Exit Sub
        Case 2
            strCaseType = DLookup("statement", "tblHICSCaseType", "ID = " & cboCaseType.Value)
            If cboCaseAction.Value = 5 Or cboCaseAction.Value = 9 Then: Exit Sub
        Case 3
            Select Case cboCaseAction.Value
                Case 1, 2, 3, 4, 6, 7, 8, 10, 11, 12, 13, 14, 15 'non-resolutions, standard resolutions, ue (unauthorized enrollment)
                    strCaseAction = DLookup("statement", "tblHICSCaseAction", "ID = " & cboCaseAction.Value)
                Case 5 'duplicate
                    txtCaseworkNotes.Value = Null
                    strCaseAction = DLookup("statement", "tblHICSCaseAction", "ID = " & cboCaseAction.Value)
                    txtCaseworkNotes = strCaseAction
                    txtCMSResolutionStatement.Value = DLookup("cms_resolution", "tblHICSLetterResponses", "case_type_id = '11' AND outcome_type_id = '5'")
                    txtResolutionOutcome.Value = DLookup("resolution_code", "tblHICSLetterResponses", "case_type_id = '11' AND outcome_type_id = '5'")
                    txtOutcome.Value = DLookup("outcome", "tblHICSLetterResponses", "case_type_id = '11' AND outcome_type_id = '" & cboCaseAction.Value & "'")
                    Call txtDuplicateCaseID_Exit(0)
                    Exit Sub
                Case 9 'transferred
                    txtCaseworkNotes.Value = Null
                    strCaseAction = DLookup("statement", "tblHICSCaseAction", "ID = " & cboCaseAction.Value)
                    txtCaseworkNotes = strCaseAction
                    Exit Sub 'null
                Case Else
                    strCaseAction = Null & vbNullString
            End Select
    End Select
    
    'get resolutions, if necassary
    If cboResolveCase.Value = -1 And (cboCaseAction.Value <> 5 And cboCaseAction.Value <> 9) Then '--REGULAR RESOLUTION--
        txtResponseSentence.Value = DLookup("letter_response", "tblHICSLetterResponses", "case_type_id = '" & cboCaseType.Value & "' AND outcome_type_id = '" & cboCaseAction.Value & "'")
        If Not IsNull(txtCaseID) And Not IsNull(txtResponseSentence) Then: txtResponseSentence.Value = Replace(txtResponseSentence.Value, "EXXXXXXXXXX", txtCaseID.Value)
        txtCMSResolutionStatement.Value = DLookup("cms_resolution", "tblHICSLetterResponses", "case_type_id = '" & cboCaseType.Value & "' AND outcome_type_id = '" & cboCaseAction.Value & "'")
        txtResolutionOutcome.Value = DLookup("resolution_code", "tblHICSLetterResponses", "case_type_id = '" & cboCaseType.Value & "' AND outcome_type_id = '" & cboCaseAction.Value & "'")
        txtOutcome.Value = DLookup("outcome", "tblHICSLetterResponses", "case_type_id = '" & cboCaseType.Value & "' AND outcome_type_id = '" & cboCaseAction.Value & "'")
    End If
    
    'send notes to form w/ pipe formatting
    txtCaseworkNotes.Value = strInteraction & vbCrLf & "|" & vbCrLf & strCaseType & vbCrLf & "|" & vbCrLf & strCaseAction

End Sub
Public Sub GetCaseActList()
    If cboResolveCase.Value = 0 And cboCaseType.Value <> 13 Then 'open
        cboCaseAction.RowSource = "SELECT ID, action, statement FROM tblHICSCaseAction WHERE type = 'open' ORDER BY ID;"
    ElseIf cboResolveCase.Value = -1 And cboCaseType.Value = 2 Then  'coe
        cboCaseAction.RowSource = "SELECT ID, action, statement FROM tblHICSCaseAction WHERE type = 'coe' ORDER BY ID;"
    ElseIf cboResolveCase.Value = 0 And cboCaseType.Value = 13 Then 'ue - further research
        cboCaseAction.RowSource = "SELECT ID, action, statement FROM tblHICSCaseAction WHERE type = 'ue - open' ORDER BY ID;"
    ElseIf cboResolveCase.Value = -1 And cboCaseType.Value = 13 Then 'ue - resolve
        cboCaseAction.RowSource = "SELECT ID, action, statement FROM tblHICSCaseAction WHERE type = 'ue - close' ORDER BY ID;"
    Else 'regular res
        cboCaseAction.RowSource = "SELECT ID, action, statement FROM tblHICSCaseAction WHERE type = 'resolution' ORDER BY ID;"
    End If
End Sub
Public Sub ClearRes()
    Deactivate txtResponseSentence
    Deactivate txtCMSResolutionStatement
    Deactivate txtResolutionOutcomeOtherText
    Deactivate txtDuplicateCaseID
    Deactivate cboTransferType
    Deactivate txtERRNoteReview
    Deactivate optAppealForm
    lblResponseSentence.Visible = False
    lblCMSResolutionStatement.Visible = False
    lblResolutionOutcomeOtherText.Visible = False
    lblDuplicateCaseID.Visible = False
    lblTransferType.Visible = False
    lblAppealForm.Visible = False
    cmdOpenDisputeForm.Visible = False
    txtResolutionOutcome = Null
    txtOutcome = Null
End Sub

Function fOSUserName() As String
' Returns the network id
    Dim lngLen As Long, lngX As Long
    Dim strUserName As String
    strUserName = String$(254, 0)
    lngLen = 255
    lngX = apiGetUserName(strUserName, lngLen)
    If lngX <> 0 Then
        fOSUserName = Left$(strUserName, lngLen - 1)
    Else
        fOSUserName = ""
    End If
End Function
Public Sub CheckDB()
    If db Is Nothing Then
        Set db = CurrentDb
    End If
End Sub

Function Clean(txtClean As String)
        txtClean = Replace(txtClean, vbLf, " ")
        txtClean = Replace(txtClean, vbTab, " ")
        txtClean = Replace(txtClean, vbCr, " ")
        txtClean = Replace(txtClean, vbCrLf, " ")
        txtClean = Replace(txtClean, vbNewLine, " ")
        txtClean = Replace(txtClean, ";", ":")
        txtClean = Replace(txtClean, "<", "(")
        txtClean = Replace(txtClean, ">", ")")
        txtClean = Replace(txtClean, Chr(160), " ")
        txtClean = Replace(txtClean, "'", " ")
        txtClean = Replace(txtClean, Chr(146), " ")
        txtClean = Replace(txtClean, Chr(39), " ")
        txtClean = Replace(txtClean, "(", "")
        txtClean = Replace(txtClean, ")", "")
        txtClean = Replace(txtClean, "-", "")
        'txtClean = Replace(txtClean, "|", "  ")
        If InStr(txtClean, "Exxxxxxxxxxx") <> "0" Then
            txtClean = Replace(txtClean, "Exxxxxxxxxxx", Me.txtCaseID.Value)
        End If
        Trim (txtClean)
        Clean = txtClean
End Function
Private Sub LockForm()
    'for searching closed cases - only txtInternalNotes is enabled
    With Me
        Dim ctrl As Control
        For Each ctrl In Controls
            If TypeOf ctrl Is TextBox And (ctrl.Name <> "txtInternalNotes" And ctrl.Name <> "txtCaseNarrative" And ctrl.Name <> "txtAddCaseNotes" _
            And ctrl.Name <> "txtERRNotes") Then
                ctrl.Enabled = False
            ElseIf TypeOf ctrl Is ComboBox Then
                ctrl.Enabled = False
            End If
        Next
    End With
End Sub
Private Sub ClearForm()
    With Me
        Dim ctrl As Control
        For Each ctrl In Controls
            If TypeOf ctrl Is TextBox Then
                ctrl.Value = Null
                ctrl.Enabled = True
            ElseIf TypeOf ctrl Is ComboBox Then
                ctrl.Value = Null
                ctrl.Enabled = True
            End If
        Next
    End With
End Sub
Sub Activate(ctrl As Control)
    With ctrl
        .Enabled = True
        .Locked = False
        .Visible = True
    End With
End Sub
Sub Deactivate(ctrl As Control)
    With ctrl
        .Enabled = False
        .Locked = True
        .Visible = False
        .Value = Null
    End With
End Sub


