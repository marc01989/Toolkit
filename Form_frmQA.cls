VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmQA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim varCaseWorkNotes As Variant
Dim varInternalNotes As Variant
Dim varERRNotes As Variant
Dim varCutlogNotes As Variant
Dim recordMax As Long, recordCounter As Long
Private Declare Function apiGetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'form for viewing notes and status of HICs case
'this also gives certain team members access to edit/delete notes as needed
'-----------------------------------------------------------------------------------------------------------------------------------------------------------

'REFERRENCES NEEDED FOR CALLS TO WINDOWS SHELL MESSAGE BOX:
    'Windows Script Host Object Model
    'WSHControllerLibrary
    
    
Private Sub txtAddCaseNotes_DblClick(Cancel As Integer)
    'view case notes in popup box
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
    'view case narrative in popup box
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
    'view demographic info in popup box
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

Private Sub Form_Load()
    'upon opening of form, check is user has access to update/delete buttons. if so, make them visible
    Dim qaPriv As Boolean: qaPriv = DLookup("HICS_UPLOAD", "tblEmployeeNames", "NetworkID = '" & Environ$("username") & "'")
    If qaPriv = True Then
        cmdUpdateCaseStatus.Visible = True
        cmdUpdateMemberDemo.Visible = True
        cmdUpdateNotes.Visible = True
        cmdDelete.Visible = True
    Else
        cmdUpdateCaseStatus.Visible = False
        cmdUpdateMemberDemo.Visible = False
        cmdUpdateNotes.Visible = False
        cmdDelete.Visible = False
    End If
    
    Me.AutoResize = True
    Me.FitToScreen = True
    DoCmd.Maximize
    txtCaseID.SetFocus
    txtCaseID.Locked = False
    txtCaseStatus.BackColor = RGB(255, 255, 255)

    Dim strRefCaseID As String
    If Not IsNull(Forms!frmQA.OpenArgs) Then
    strRefCaseID = Forms!frmQA.OpenArgs
        If Len(strRefCaseID) > 0 Then
            txtCaseID.Value = strRefCaseID
            cmdSearch_Click
        End If
    End If

End Sub

Private Sub cmdClear_Click()
    'clear and reset form
    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            ctrl.Value = Null
        End If
        If TypeName(ctrl) = "CheckBox" Then
            ctrl.Value = False
        End If
    Next
    varCaseWorkNotes = Empty
    varInternalNotes = Empty
    varERRNotes = Empty
    varCutlogNotes = Empty
    recordMax = 0
    cboReopenReason.Value = Null
    txtErrorBar.BackColor = RGB(245, 246, 246)
    txtErrorBox.BackColor = RGB(245, 246, 246)
    txtCaseID.SetFocus
    txtCaseID.Locked = False
    txtCaseStatus.BackColor = RGB(255, 255, 255)
    Deactivate txtDuplicateCaseID
    Deactivate txtHelpDeskDate
    Deactivate txtERRDate
    Deactivate txtERRReturnDate
    Deactivate txtTransferredTo
    Deactivate txtScanUploadDate
    Deactivate txtReopenDate
    Deactivate cboReopenReason
    Deactivate txtPermNote

End Sub


Private Sub cmdSearch_Click()
    'search button. this module takes all results and saves into local arrays, which can then be iterated through
    
    'VALIDATION
    If CheckSubmission(1) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            txtErrorBar.SetFocus
            Exit Sub
        End If
    End If

On Error GoTo err1:
    Dim db As Database: Set db = CurrentDb
    Dim rs As Recordset
    Dim regexStr As String, strTemp As String
    txtCaseID.Locked = True

        '################## CASE STATUS ###############################
        Set rs = CurrentDb.OpenRecordset("tblCaseStatus", dbOpenSnapshot)
        With rs
            .FindFirst "case_id = '" & txtCaseID.Value & "'"
            If Not .NoMatch Then
                txtResolveCase.Value = ![resolve_case]
                txtPrinted.Value = ![printed]
                txtHelpDesk.Value = ![pending_helpdesk]
                txtHelpDeskDate.Value = ![date_referred]
                txtERR.Value = ![Err]
                txtERRDate.Value = ![err_date]
                txtERRReturnDate.Value = ![err_returned]
                chkDuplicate.Value = ![duplicate]
                txtDuplicateCaseID.Value = ![duplicate_caseID]
                chkTransferred.Value = ![transferred]
                txtTransferredTo.Value = ![trasnfered_to]
                chkScanUpload.Value = ![scan_uploaded]
                txtScanUploadDate.Value = ![scan_upload_date]
                chkReopened.Value = ![reopened]
                txtReopenDate.Value = ![reopen_date]
                cboReopenReason.Value = ![reopen_reason]
                chkUpdateAllowed.Value = ![updated_allowed]
                txtPermNote.Value = ![perm_note]
                If ![resolve_case] = "Y" Then
                    txtCaseStatus.Value = "Closed"
                    txtCaseStatus.BackColor = RGB(255, 99, 71)
                ElseIf ![resolve_case] = "N" Then
                    txtCaseStatus.Value = "Open"
                    txtCaseStatus.BackColor = RGB(144, 238, 144)
                    msgStr = "This case has not been resolved"
                End If
            End If
        End With
        rs.Close
        
        ' ###################### CMS CASE DATA #######################################
        Set rs = db.OpenRecordset("tblCMSCaseData", dbOpenSnapshot)
        With rs
            .FindFirst "case_id = '" & txtCaseID.Value & "'"
            If Not .NoMatch Then
                txtComplianceDueDate.Value = Format(![compliance_due_date], "m/d/yyyy")
                txtIssuerAssignDate.Value = Format(![issuer_assign_date], "m/d/yyyy")
                txtCaseLevel.Value = ![case_level]
                txtCongressionalCase.Value = ![congressional_case]
                If Not IsNull(![case_narrative]) And Not IsNull(![Subcategory]) Then
                    regexStr = RegexSwitch(![Subcategory])
                    txtCaseNarrative.Value = RegexNarrative(![case_narrative], regexStr)
                End If
            End If
        End With
        rs.Close
        
        '####################### COMMENTS FROM HICS SYSTEM ######################
        Set rs = db.OpenRecordset("tblComments", dbOpenSnapshot)
        With rs
            .FindFirst "case_id = '" & txtCaseID.Value & "'"
            Do While .EOF = False
                If Not .NoMatch Then
                    strTemp = "<b>Type: </b>" & ![comment_type] & "<br>" & _
                    "<b>Date: </b>" & ![created_date] & "<br>" & _
                    "<b>Comment: </b>" & Trim(![comment]) & "<br><br>" & strTemp
                ElseIf .NoMatch Then
                    .Close
                    Exit Do
                End If
                .FindNext "case_id = '" & txtCaseID.Value & "'"
            Loop
        End With
        txtAddCaseNotes.Value = strTemp
        
        '###################### CONTACT INFO FROM HICS #############################
        Set rs = db.OpenRecordset("tblContact", dbOpenSnapshot)
        With rs
            .FindFirst "case_id = '" & txtCaseID.Value & "'"
            If Not .NoMatch Then
                txtDemographic.Value = _
                    "<b>Name: </b>" & ![first_name] & " " & ![last_name] & "<br><br>" & _
                    "<b>Address: </b>" & ![Address] & " " & ![address_2] & "<br>" & _
                            ![city] & ", " & ![State] & " " & ![zip_code] & "<br><br>" & _
                    "<b>E-mail: </b>" & ![email] & "<br><br>" & _
                    "<b>Phone: </b>" & ![phone] & "<br><br>" & _
                    "<b>Alt-Phone: </b>" & ![alt_phone] & "<br><br>" & _
                    "<b>DOB: </b>" & ![dob]
            End If
        End With
        rs.Close
    
        '################ DEMOGRAPHICS #################
        Set rs = CurrentDb.OpenRecordset("tblMemberDemo", dbOpenSnapshot)
        With rs
            .FindFirst "case_id = '" & txtCaseID.Value & "'"
            If Not .NoMatch Then
                txtMemberNumber.Value = ![member_id]
                txtFirstName.Value = ![first_name]
                txtLastName.Value = ![last_name]
                txtStreetAddress.Value = ![street_address]
                txtStreetAddress2.Value = ![street_address_2]
                txtCity.Value = ![city]
                txtState.Value = ![State]
                txtZIpCode.Value = ![zip_code]
            End If
        End With
        rs.Close
        
        '################ CASEWORK NOTES #################
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblCaseWorkNotes WHERE case_id = '" & txtCaseID.Value & "' ORDER BY submit_date;", dbOpenSnapshot)
        With rs
            If .recordCount > 0 Then
                .MoveFirst
                .MoveLast
                .MoveFirst
                recordMax = .recordCount - 1
                recordCounter = 0
                varCaseWorkNotes = .GetRows(.recordCount)
            End If
        End With
        rs.Close
        

        '################ INTERNAL NOTES #################
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblInternalNotes WHERE case_id = '" & txtCaseID.Value & "' ORDER BY submit_date;", dbOpenSnapshot)
        With rs
            If .recordCount > 0 Then
                varInternalNotes = .GetRows(.recordCount)
            End If
        End With
        rs.Close
        
        
        '################ ERR NOTES #################
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblERRNotes WHERE CaseID = '" & txtCaseID.Value & "' ORDER BY SubmitDate;", dbOpenSnapshot)
        With rs
            If .recordCount > 0 Then
                varERRNotes = .GetRows(.recordCount)
                
            End If
        End With
        rs.Close
        
        
        '################ CUTLOG NOTES #################
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblCutlogNotes WHERE case_id = '" & txtCaseID.Value & "' ORDER BY submit_date;", dbOpenSnapshot)
        With rs
            If .recordCount > 0 Then
                varCutlogNotes = .GetRows(.recordCount)
                
            End If
        End With
        rs.Close: Set rs = Nothing
        db.Close: Set db = Nothing
      
    Call txthelpdesk_AfterUpdate
    Call txterr_AfterUpdate
    Call chkDuplicate_AfterUpdate
    Call chktransferred_AfterUpdate
    Call chkScanUpload_AfterUpdate
    Call chkReopened_AfterUpdate
    Call chkUpdateAllowed_AfterUpdate
    
    'after all data is queried and saved into arrays - test if there are results. if so, call sub to populate the boxes on form
    If Not IsEmpty(varCaseWorkNotes) Then
        Call populateBoxes(recordCounter)
    Else
        MsgBox ("No casework notes found")
    End If
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmQA; cmdSearch_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
End Sub

Public Sub populateBoxes(i As Long)

Dim intNotes As String

On Error GoTo err1:
    'populate fields on form with results of search query

    txtNoteId.Value = varCaseWorkNotes(15, i)
    txtCaseworkNotes.Value = Replace(varCaseWorkNotes(1, i), "|", vbCrLf & vbCrLf)
    txtSubmitDate.Value = varCaseWorkNotes(2, i)
    txtNoteUploaded.Value = varCaseWorkNotes(4, i)
    If Not IsNull(varCaseWorkNotes(12, i)) Then: _
        txtSpecialist.Value = DLookup("specialist_name", "tblEmployeeNames", "employee_id = " & CInt(varCaseWorkNotes(12, i)))
    txtResolution.Value = varCaseWorkNotes(6, i)
    txtLetterResponse.Value = varCaseWorkNotes(7, i)
    If Not IsNull(varCaseWorkNotes(13, i)) Then: _
        txtInteractionType.Value = DLookup("contact_type", "tblHICSInteractionType", "ID = " & CInt(varCaseWorkNotes(13, i)))
    If Not IsNull(varCaseWorkNotes(14, i)) Then: _
        txtCaseAction.Value = DLookup("action", "tblHICSCaseAction", "ID = " & CInt(varCaseWorkNotes(14, i)))
    txtRecordCounter.Value = i + 1
    txtRecordMax.Value = recordMax + 1

    'FIND MATCHING INTERNAL NOTE(s)
    If Not IsEmpty(varInternalNotes) Then
        For j = 0 To UBound(varInternalNotes, 2)
            If varInternalNotes(5, j) = varCaseWorkNotes(15, i) Then
                'txtInternalNote.Value = varInternalNotes(2, j)
                intNotes = intNotes & varInternalNotes(3, j) & vbCrLf & varInternalNotes(2, j) & vbCrLf & vbCrLf
                txtInternalNoteId.Value = varInternalNotes(0, j)
                'Exit For
            Else
                txtInternalNote.Value = Null
                txtInternalNoteId.Value = Null
            End If
        Next j
        txtInternalNote = intNotes
    End If
    
    'FIND MATCHING ERR NOTE
    If Not IsEmpty(varERRNotes) Then
        For j = 0 To UBound(varERRNotes, 2)
            If varERRNotes(8, j) = varCaseWorkNotes(15, i) Then
                txtERRNote.Value = varERRNotes(2, j)
                txtERRNoteId.Value = varERRNotes(0, j)
                Exit For
            Else
                txtERRNote.Value = Null
                txtERRNoteId.Value = Null
            End If
        Next j
    End If
    
    'FIND MATCHING CUTLOG NOTE
    If Not IsEmpty(varCutlogNotes) Then
        For j = 0 To UBound(varCutlogNotes, 2)
            If varCutlogNotes(9, j) = varCaseWorkNotes(15, i) Then
                txtCutlogNote.Value = varCutlogNotes(2, j)
                txtCutlogNoteId.Value = varCutlogNotes(8, j)
                Exit For
            Else
                txtCutlogNote.Value = Null
                txtCutlogNoteId.Value = Null
            End If
        Next j
    End If

err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmQA; PopulateBoxes()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select

End Sub
Private Sub cmdPrevNote_Click()
    'previous record
    If recordMax = 0 Then Exit Sub
    If recordCounter = 0 Then
        Exit Sub
    Else
        recordCounter = recordCounter - 1
        Call populateBoxes(recordCounter)
    End If
End Sub
Private Sub cmdNextNote_Click()
    'next record
    If recordMax = 0 Then Exit Sub
    If recordCounter = recordMax Then
        Exit Sub
    Else
        recordCounter = recordCounter + 1
        Call populateBoxes(recordCounter)
    End If
End Sub


Private Sub cmdUpdateCaseStatus_Click()
'update the status of a case. status of a case encompasses a bunch of different data points. this data lives on tblCaseStatus

    'VALIDATION
    If CheckSubmission(2) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            txtErrorBar.SetFocus
            Exit Sub
        End If
    End If
    
On Error GoTo err1:
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Dim networkID As String: networkID = Environ$("username")
    
    'update record with values from form
    Set rs = db.OpenRecordset("tblCaseStatus", dbOpenDynaset)
    With rs
        .FindFirst "case_id = '" & txtCaseID.Value & "'"
        If Not .NoMatch Then
            If MsgBox("Update case status info?", vbQuestion + vbYesNo, "Update record?") = vbYes Then
                .Edit
                    ![resolve_case] = txtResolveCase
                    ![printed] = txtPrinted
                    ![pending_helpdesk] = txtHelpDesk
                    ![date_referred] = txtHelpDeskDate
                    ![Err] = txtERR
                    ![err_date] = txtERRDate
                    ![err_returned] = txtERRReturnDate
                    ![duplicate] = chkDuplicate
                    ![duplicate_caseID] = txtDuplicateCaseID
                    ![transferred] = chkTransferred
                    ![trasnfered_to] = txtTransferredTo
                    ![scan_uploaded] = chkScanUpload
                    ![scan_upload_date] = txtScanUploadDate
                    ![reopened] = chkReopened
                    ![reopen_date] = txtReopenDate
                    ![reopen_reason] = cboReopenReason
                .Update
                .Close
                MsgBox ("Case Status Info updated.")
            End If
        Else
        
        End If
    End With
    
    Set rs = Nothing
    db.Close: Set db = Nothing
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmQA; cmdUpdateCaseStatus_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select

End Sub

Private Sub cmdUpdateMemberDemo_Click()
'update demographic info

    If CheckSubmission(3) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            txtErrorBar.SetFocus
            Exit Sub
        End If
    End If
    
On Error GoTo err1:
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Dim networkID As String: networkID = Environ$("username")
    
    Set rs = db.OpenRecordset("tblMemberDemo", dbOpenDynaset)
    With rs
        .FindFirst "case_id = '" & txtCaseID.Value & "'"
        If Not .NoMatch Then
            'update existing record
            If MsgBox("Update demographic info?", vbQuestion + vbYesNo, "Update record?") = vbYes Then
                .Edit
                    ![member_id] = txtMemberNumber
                    ![first_name] = txtFirstName
                    ![last_name] = txtLastName
                    ![street_address] = txtStreetAddress
                    ![street_address_2] = txtStreetAddress2
                    ![city] = txtCity
                    ![State] = txtState
                    ![zip_code] = txtZIpCode
                    ![employee_id] = DLookup("employee_id", "tblEmployeeNames", "NetworkID = '" & networkID & "'")
                .Update
                .Close
                MsgBox ("Demographic Info updated.")
            End If
        Else
            'add new record
            If MsgBox("Add new demographic record for case " & txtCaseID & "?", vbQuestion + vbYesNo, "Add new record?") = vbYes Then
                .AddNew
                    ![case_id] = txtCaseID
                    ![member_id] = txtMemberNumber
                    ![first_name] = txtFirstName
                    ![last_name] = txtLastName
                    ![street_address] = txtStreetAddress
                    ![street_address_2] = txtStreetAddress2
                    ![city] = txtCity
                    ![State] = txtState
                    ![zip_code] = txtZIpCode
                    ![employee_id] = DLookup("employee_id", "tblEmployeeNames", "NetworkID = '" & networkID & "'")
                .Update
                .Close
                MsgBox ("New Demographic Info added.")
            End If
        End If
    End With
    
    Set rs = Nothing
    db.Close: Set db = Nothing
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmQA; cmdUpdateMemberDemo_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
     
End Sub
Private Sub cmdUpdateNotes_Click()

'update casework notes and all corresponding notes (internal, cutlog, err note)

    If CheckSubmission(4) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            txtErrorBar.SetFocus
            Exit Sub
        End If
    End If
    
On Error GoTo err1:
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Dim networkID As String: networkID = Environ$("username")
    Dim msgStr As String
    
    '################ CASEWORK NOTES #################
    Set rs = db.OpenRecordset("tblCaseWorkNotes", dbOpenDynaset)
    With rs
        .FindFirst "ID = " & txtNoteId.Value
        If Not .NoMatch Then
        
            If MsgBox("Update case notes?", vbQuestion + vbYesNo, "Update Notes?") = vbYes Then
                .Edit
                ![casework_notes] = Clean(txtCaseworkNotes)
                If Not IsNull(varCaseWorkNotes(7, recordCounter)) Then: ![resolution_response] = txtLetterResponse
                .Update
                .Close
                
                varCaseWorkNotes(1, recordCounter) = txtCaseworkNotes
                If Not IsNull(varCaseWorkNotes(7, recordCounter)) Then: varCaseWorkNotes(7, recordCounter) = txtLetterResponse
            Else
                Exit Sub
            End If
        Else
            msgStr = msgStr & "No casework note found." & vbCrLf
        End If
    End With
    
    '################ INTERNAL NOTES #################
    Set rs = db.OpenRecordset("tblInternalNotes", dbOpenDynaset)
    With rs
        .FindFirst "casework_note_id = " & txtNoteId.Value
        If Not .NoMatch Then
            .Edit
            ![Note] = txtInternalNote
            .Update
            .Close
            
            If Not IsEmpty(varInternalNotes) Then
                For j = 0 To UBound(varInternalNotes, 2)
                    If varInternalNotes(0, j) = txtInternalNoteId.Value Then
                        varInternalNotes(2, j) = txtInternalNote.Value
                        Exit For
                    End If
                Next j
            End If
        Else
            msgStr = msgStr & "No internal note found." & vbCrLf
        End If
    End With
    
    '################ ERR NOTES #################
    Set rs = db.OpenRecordset("tblERRNotes", dbOpenDynaset)
    With rs
        .FindFirst "casework_note_id = " & txtNoteId.Value
        If Not .NoMatch Then
            .Edit
            ![Note] = txtERRNote
            .Update
            .Close
            
            If Not IsEmpty(varERRNotes) Then
                For j = 0 To UBound(varERRNotes, 2)
                    If varERRNotes(0, j) = txtERRNoteId.Value Then
                        varERRNotes(2, j) = txtERRNote.Value
                        Exit For
                    End If
                Next j
            End If
        Else
            msgStr = msgStr & "No ER&&R note found." & vbCrLf
        End If
    End With
    
    '################ CUTLOG NOTES #################
    Set rs = db.OpenRecordset("tblCutlogNotes", dbOpenDynaset)
    With rs
        .FindFirst "casework_note_id = " & txtNoteId.Value
        If Not .NoMatch Then
            .Edit
            ![cutlog_notes] = txtCutlogNote
            .Update
            .Close
            
            If Not IsEmpty(varCutlogNotes) Then
                For j = 0 To UBound(varCutlogNotes, 2)
                    If varCutlogNotes(8, j) = txtCutlogNoteId.Value Then
                        varCutlogNotes(2, j) = txtCutlogNote.Value
                        Exit For
                    End If
                Next j
            End If
        Else
            msgStr = msgStr & "No ER&&R note found." & vbCrLf
        End If
    End With
    
    
    Set rs = Nothing
    db.Close: Set db = Nothing

'    If msgStr <> "" Then
'        MsgBox (msgStr)
'    Else
'        MsgBox ("Notes updated")
'    End If
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmQA; cmdUpdateMemberDemo_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub
Private Sub cmdDelete_Click()
'delete case work note and all corresponding notes (internal, cutlog, err note)

    'VALIDATION
    If CheckSubmission(5) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            txtErrorBar.SetFocus
            Exit Sub
        End If
    End If
    
On Error GoTo err1:
    Dim resolution As String: resolution = varCaseWorkNotes(6, recordCounter)
    Dim caseID As String: caseID = varCaseWorkNotes(0, recordCounter)
    Dim recordId As Long: recordId = varCaseWorkNotes(15, recordCounter)
    Dim searchStr As String: searchStr = "case_id = '" & varCaseWorkNotes(0, recordCounter) & "' AND ID = " & varCaseWorkNotes(15, recordCounter) & " AND uploaded = 'N'"
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Select Case resolution
        Case "Y"  'RESOLVED CASE - DELETE RESOLUTION NOTE AND REOPEN CASE
            If MsgBox("You are about to delete a Resolution Note - this will reopen the case." & vbCrLf & vbCrLf & "Continue?", vbQuestion + vbYesNoCancel, "Reopen Case?") = vbYes Then
                Set db = CurrentDb
                Set rs = db.OpenRecordset("tblCaseWorkNotes", dbOpenDynaset)
                rs.FindFirst searchStr
                If Not rs.NoMatch Then
                                
                    'remove closing casework note
                    With rs
                        .Delete
                        .Close
                    End With
                    
                    'update status and other fields, if necessary
                    Set rs = db.OpenRecordset("tblCaseStatus")
                        With rs
                            .FindFirst "case_id = '" & caseID & "'"
                            If Not .NoMatch Then
                            .Edit
                                ![resolve_case] = "N"
                                If ![duplicate] = "Yes" Then
                                    ![duplicate] = "No"
                                    ![duplicate_caseID] = Null
                                End If
                                If ![transferred] = -1 Then
                                    ![transferred] = 0
                                    ![trasnfered_to] = Null
                                End If
                            .Update
                            End If
                        End With
                    
                    'delete closing cutlog
                    Set rs = db.OpenRecordset("tblCutlogNotes")
                        rs.FindFirst "casework_note_id = " & recordId
                        If Not rs.NoMatch Then
                            With rs
                                .Delete
                                .Close
                            End With
                        End If
                        
                    'delete internal note
                    Set rs = db.OpenRecordset("tblInternalNotes")
                        rs.FindFirst "casework_note_id = " & recordId
                        If Not rs.NoMatch Then
                            With rs
                                .Delete
                                .Close
                            End With
                        End If
                    
                    'reload form to show case is reopened
                    Call cmdClear_Click
                    txtCaseID.Value = caseID
                    Call cmdSearch_Click
                Else
                    MsgBox ("Cannot delete record")
                    Exit Sub
                End If
            Else: Exit Sub
            End If
            
            
        Case "N"  'OPEN CASE - DELETE NOTE
            If MsgBox("Do you want to delete this casework note?", vbQuestion + vbYesNoCancel, "Delete Note?") = vbYes Then
                Set db = CurrentDb
                Set rs = db.OpenRecordset("tblCaseWorkNotes", dbOpenDynaset)
                rs.FindFirst searchStr
                If Not rs.NoMatch Then
                    With rs
                        .Delete
                        .Close
                    End With
                    
                    'delete cutlog
                    Set rs = db.OpenRecordset("tblCutlogNotes")
                        rs.FindFirst "casework_note_id = " & recordId
                        If Not rs.NoMatch Then
                            With rs
                                .Delete
                                .Close
                            End With
                        End If
                        
                    'delete internal note
                    Set rs = db.OpenRecordset("tblInternalNotes")
                        rs.FindFirst "casework_note_id = " & recordId
                        If Not rs.NoMatch Then
                            With rs
                                .Delete
                                .Close
                            End With
                        End If
                    
                    'delete err note
                    Set rs = db.OpenRecordset("tblERRNotes")
                        rs.FindFirst "casework_note_id = " & recordId
                        If Not rs.NoMatch Then
                            With rs
                                .Delete
                                .Close
                            End With
                            
                            'reset values for ERR data
                            Set rs = db.OpenRecordset("tblCaseStatus")
                            rs.FindFirst "case_id = '" & caseID & "'"
                            If Not rs.NoMatch Then
                                With rs
                                    .Edit
                                    ![Err] = "N"
                                    ![err_date] = Null
                                    .Update
                                End With
                            End If
                        End If
                    
                End If
                Call cmdClear_Click
                txtCaseID.Value = caseID
                Call cmdSearch_Click
            Else: Exit Sub
            End If

        Case Else: Exit Sub
    
    'CLEAN UP
    End Select
    If Not rs Is Nothing Then
        Set rs = Nothing
    End If
    If Not db Is Nothing Then
        Set db = Nothing
    End If
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmQA; cmdDelete_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
    
    
End Sub

Private Sub txtERR_Exit(Cancel As Integer)
    If (txtERR.Value <> "Y" And txtERR.Value <> "N") Then
        If MsgBox("Value must be 'Y' or 'N'.", vbExclamation + vbOKOnly, "Error") = vbOK Then
            Cancel = True
        End If
    Else
        txtERR.Value = UCase(txtERR)
    End If
End Sub

Private Sub txtHelpDesk_Exit(Cancel As Integer)
    If (txtHelpDesk.Value <> "Y" And txtHelpDesk.Value <> "N") Then
        If MsgBox("Value must be 'Y' or 'N'.", vbExclamation + vbOKOnly, "Error") = vbOK Then
            Cancel = True
        End If
    Else
        txtHelpDesk.Value = UCase(txtHelpDesk)
    End If
End Sub

Private Sub txtPrinted_Exit(Cancel As Integer)
    If (txtPrinted.Value <> "Y" And txtPrinted.Value <> "N") Then
        If MsgBox("Value must be 'Y' or 'N'.", vbExclamation + vbOKOnly, "Error") = vbOK Then
            Cancel = True
        End If
    Else
        txtPrinted.Value = UCase(txtPrinted)
    End If
End Sub

Private Sub txtResolveCase_Exit(Cancel As Integer)
    If (txtResolveCase.Value <> "Y" And txtResolveCase.Value <> "N") Then
        If MsgBox("Value must be 'Y' or 'N'.", vbExclamation + vbOKOnly, "Error") = vbOK Then
            Cancel = True
        End If
    Else
        txtResolveCase.Value = UCase(txtResolveCase)
    End If
End Sub
Private Sub txthelpdesk_AfterUpdate()
    If txtHelpDesk.Value = "Y" Then
        Activate txtHelpDeskDate
    Else
        Deactivate txtHelpDeskDate
    End If
End Sub
Private Sub txterr_AfterUpdate()
    If txtERR.Value = "Y" Then
        Activate txtERRDate
        Activate txtERRReturnDate
    ElseIf txtERR.Value = "N" And Not IsNull(txtERRDate) Then
        Activate txtERRDate
        Activate txtERRReturnDate
    Else
        Deactivate txtERRDate
        Deactivate txtERRReturnDate
    End If
End Sub
Private Sub chkDuplicate_AfterUpdate()
    If chkDuplicate.Value = True Then
        Activate txtDuplicateCaseID
    Else
        Deactivate txtDuplicateCaseID
    End If
End Sub
Private Sub chktransferred_AfterUpdate()
    If chkTransferred.Value = True Then
        Activate txtTransferredTo
    Else
        Deactivate txtTransferredTo
    End If
End Sub
Private Sub chkScanUpload_AfterUpdate()
    If chkScanUpload.Value = True Then
        Activate txtScanUploadDate
    Else
        Deactivate txtScanUploadDate
    End If
End Sub
Private Sub chkReopened_AfterUpdate()
    If chkReopened.Value = True Then
        Activate txtReopenDate
        Activate cboReopenReason
    Else
        Deactivate txtReopenDate
        Deactivate cboReopenReason
    End If
End Sub
Private Sub chkUpdateAllowed_AfterUpdate()
    If chkUpdateAllowed.Value = True Then
        Activate txtPermNote
    Else
        Deactivate txtPermNote
    End If
End Sub
Public Function CheckSubmission(submitType As Integer) As Boolean

    Dim msgStr As String, headerStr As String, footerStr As String, resultStr As String
    Dim regex As RegExp: Set regex = New RegExp
    headerStr = "<ul>"
    footerStr = "</ul>"
    
    Select Case submitType
        Case 1 '--SEARCH BUTTON--
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
        Case 2 '--UPDATE CASE STATUS --

            With regex
                .MultiLine = False
                .Global = True
                .IgnoreCase = True
                .Pattern = "^(0?[1-9]|1[0-2])\/(0?[1-9]|1[0-9]|2[0-9]|3[01])\/\d{4}$"
            End With
            'resolve case
            If IsNull(txtResolveCase) Or (txtResolveCase.Value <> "N" And txtResolveCase.Value <> "Y") Then
                msgStr = msgStr & "<li><b>Resolve case</b> must be 'Y' or 'N'.</li>"
            End If
            'printed
            If IsNull(txtPrinted) Or (txtPrinted.Value <> "N" And txtPrinted.Value <> "Y") Then
                msgStr = msgStr & "<li><b>Printed</b> must be 'Y' or 'N'.</li>"
            End If
            'helpdesk
            If IsNull(txtHelpDesk) Or (txtHelpDesk.Value <> "N" And txtHelpDesk.Value <> "Y") Then
                msgStr = msgStr & "<li><b>Helpdesk</b> must be 'Y' or 'N'.</li>"
            End If
            If txtHelpDesk.Value = "Y" And IsNull(txtHelpDeskDate.Value) Then
                msgStr = msgStr & "<li><b>Helpdesk Date</b> cannot be blank.</li>"
            End If
            If Not IsNull(txtHelpDeskDate) Then
                If regex.test(txtHelpDeskDate.Value) = False Then
                    msgStr = msgStr & "<li><b>Helpdesk Date</b> must be 'mm/dd/yyyy' format.</li>"
                End If
            End If
            'err
            If IsNull(txtERR) Or (txtERR.Value <> "N" And txtERR.Value <> "Y") Then
                msgStr = msgStr & "<li><b>ERR</b> must be 'Y' or 'N'.</li>"
            End If
            If txtERR.Value = "Y" And IsNull(txtERRDate.Value) Then
                msgStr = msgStr & "<li><b>ERR Date</b> cannot be blank.</li>"
            End If
            If txtERR.Value = "N" And ((IsNull(txtERRDate.Value) And Not IsNull(txtERRReturnDate)) Or (Not IsNull(txtERRDate.Value) And IsNull(txtERRReturnDate))) Then
                msgStr = msgStr & "<li><b>ERR</b> is set to No: please check that date boxes both have values or are both empty.</li>"
            End If
            If Not IsNull(txtERRDate) Then
                If regex.test(txtERRDate.Value) = False Then
                    msgStr = msgStr & "<li><b>ERR Date</b> must be 'mm/dd/yyyy' format.</li>"
                End If
            End If
            If Not IsNull(txtERRReturnDate) Then
                If regex.test(txtERRReturnDate.Value) = False Then
                    msgStr = msgStr & "<li><b>ERR Return Date</b> must be 'mm/dd/yyyy' format.</li>"
                End If
            End If
            'duplicate
            If chkDuplicate.Value = True And IsNull(txtDuplicateCaseID.Value) Then
                msgStr = msgStr & "<li><b>Duplicate Case ID</b> cannot be blank.</li>"
            End If
            regex.Pattern = "^[eE0-9]{11}$"
            If Not IsNull(txtDuplicateCaseID) Then
                If regex.test(txtDuplicateCaseID.Value) = False Then
                    msgStr = msgStr & "<li><b>Duplicate Case ID</b> is invalid.</li>"
                End If
            End If
            'transferred
            If chkTransferred.Value = True And IsNull(txtTransferredTo.Value) Then
                msgStr = msgStr & "<li><b>Transferred To</b> cannot be blank.</li>"
            End If
            'scans
            If chkScanUpload.Value = True And IsNull(txtScanUploadDate.Value) Then
                msgStr = msgStr & "<li><b>Scan Upload Date</b> cannot be blank.</li>"
            End If
            regex.Pattern = "^((0?[1-9]|1[0-2])\/(0?[1-9]|1[0-9]|2[0-9]|3[01])\/\d{4})( ([0-9]|1[0-2]):(0[0-9]|[1-59]\d):(0[0-9]|[1-59]\d)\s(AM|am|PM|pm))?$"
            If Not IsNull(txtScanUploadDate) Then
                If regex.test(txtScanUploadDate.Value) = False Then
                    msgStr = msgStr & "<li><b>Scan Upload Date</b> must be 'mm/dd/yyyy' format.</li>"
                End If
            End If


            'reopened
            If chkReopened.Value = True And IsNull(txtReopenDate.Value) Then
                msgStr = msgStr & "<li><b>Reopen Date</b> cannot be blank.</li>"
            End If
            If chkReopened.Value = True And IsNull(cboReopenReason.Value) Then
                msgStr = msgStr & "<li><b>Reopen Reason</b> cannot be blank.</li>"
            End If
            regex.Pattern = "^(0?[1-9]|1[0-2])\/(0?[1-9]|1[0-9]|2[0-9]|3[01])\/\d{4}$"
            If Not IsNull(txtReopenDate) Then
                If regex.test(txtReopenDate.Value) = False Then
                    msgStr = msgStr & "<li><b>Reopen Date</b> must be 'mm/dd/yyyy' format.</li>"
                End If
            End If
            
        Case 3 'UPDATE MEMBER DEMO
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
            txtMemberNumber.Value = ScrubMemberId(txtMemberNumber.Value & vbNullString)
            If txtMemberNumber.Value = "" Or IsNull(txtMemberNumber.Value) = True Then
                msgStr = msgStr & "<li><b>Member Number</b> cannot be blank.</li>"
            End If
            regex.Pattern = "^[0-9]{11}$"
            If Not IsNull(txtMemberNumber) Then
                If regex.test(txtMemberNumber.Value) = False Then
                    msgStr = msgStr & "<li><b>Member Number</b> is invalid. Please try again.</li>"
                End If
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
            
        Case 4 'UPDATE CASEWORK NOTES
            If IsEmpty(varCaseWorkNotes) Then
                msgStr = msgStr & "<li>You must select a casework note to update.</li>"
            Else
                If Not IsNull(txtLetterResponse) And txtResolution.Value = "N" Then
                    msgStr = msgStr & "<li>Cannot add <b>Letter Response</b> to non-resolution note - please remove all data from this box</li>"
                End If
                If Not IsNull(txtInternalNote) And IsNull(DLookup("case_id", "tblInternalNotes", "casework_note_id = " & varCaseWorkNotes(15, recordCounter))) Then
                    msgStr = msgStr & "<li>No existing <b>Internal Note</b> to modify - please remove all data from this box.</li>"
                End If
                If Not IsNull(txtERRNote) And IsNull(DLookup("CaseID", "tblERRNotes", "casework_note_id = " & varCaseWorkNotes(15, recordCounter))) Then
                    msgStr = msgStr & "<li>No existing <b>ER&R Note</b> to modify - please remove all data from this box.</li>"
                End If
                If Not IsNull(txtCutlogNote) And IsNull(DLookup("case_id", "tblCutlogNotes", "casework_note_id = " & varCaseWorkNotes(15, recordCounter))) Then
                    msgStr = msgStr & "<li>No existing <b>Cutlog Note</b> to modify - please remove all data from this box.</li>"
                End If
            End If
            
        Case 5 'DELETE CASEWORK NOTES
            If IsEmpty(varCaseWorkNotes) Then
                msgStr = msgStr & "<li>You must select a casework note to delete.</li>"
            End If
            If txtNoteUploaded.Value = "Y" Then
                msgStr = msgStr & "<li>Cannot delete record - casework note has already been loaded to CMS system..</li>"
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
