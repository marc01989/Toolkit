VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Record Locator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim recordMax As Long, recordCounter As Long
Dim arrayErrors As Variant
Dim userId As Integer

Private Sub Form_Load()
    Dim searchVal As Variant, passedIDs As Variant
    Dim obj As AccessObject, dbs As Object
    
    'checks if value is passed when opening form, runs main query if so
    If Not IsNull(Forms![Record Locator].OpenArgs) Then
        searchVal = Forms![Record Locator].OpenArgs
        If Len(searchVal) > 0 Then
            
            passedIDs = Split(searchVal, ",")
            
                If passedIDs(2) = 0 Or IsNull(passedIDs) Then
                    MsgBox ("User not found - please login again")
                    Set dbs = Application.CurrentProject
                    For Each obj In dbs.AllForms
                        If obj.IsLoaded = True Then
                          DoCmd.Close acForm, obj.Name, acSaveNo
                        End If
                    Next obj
                    DoCmd.OpenForm "Login", acNormal, , , , acWindowNormal
                    Call LogError(0 & " " & "User Id not found or not passed to home screen", "Home; Form_Load()")
                    Exit Sub
                    
                Else
                    userId = CInt(passedIDs(2))
                    'MsgBox (userId)
                End If
                
            If passedIDs(0) = "" Or passedIDs(1) = "" Then Exit Sub
            cboWeekOf.Value = CInt(Trim(passedIDs(0)))
            cboSpecialist.Value = CInt(Trim(passedIDs(1)))
            btnSearch_Click
        End If
    Else
        MsgBox ("User not found - please login again")
        Set dbs = Application.CurrentProject
        For Each obj In dbs.AllForms
            If obj.IsLoaded = True Then
              DoCmd.Close acForm, obj.Name, acSaveNo
            End If
        Next obj
        DoCmd.OpenForm "Login", acNormal, , , , acWindowNormal
        Call LogError(0 & " " & "User Id not found or not passed to home screen", "Home; Form_Load()")
        Exit Sub
    End If

End Sub
Private Sub btnSearch_Click()
    
    If ValidateForm(1) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If
    
On Error GoTo err1:
    'runs sql, selects all qa'd items where matches on week and employee
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM review_items_mkt WHERE (week_id = " & cboWeekOf.Value & " AND employee_id = " & cboSpecialist.Value & ") ORDER BY ID;", dbOpenSnapshot, dbReadOnly)
    With rs
        If .recordCount > 0 Then
            .MoveFirst
            .MoveLast
            .MoveFirst
            'save into array, set recordMax and recordCounter
            recordMax = .recordCount - 1
            recordCounter = 0
            txtRecordCount.Value = .recordCount
            arrayErrors = .GetRows(.recordCount)
            Call populateBoxes(recordCounter)
        Else
            MsgBox "No records found"
            Call ClearForm
            arrayChats = Empty
            recordMax = 0
            recordCounter = 0
        End If
    End With
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing

err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Record Locator; btnSearch_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
        
End Sub
Public Sub populateBoxes(i As Long)

    txtID.Value = arrayErrors(0, i)
    txtMemberID.Value = arrayErrors(1, i)
    txtAdditionalId.Value = arrayErrors(2, i)
    txtReportName.Value = arrayErrors(3, i)
    txtErrorComment.Value = arrayErrors(4, i)
    txtSubmitDate.Value = arrayErrors(5, i)
    cboWeekOf.Value = arrayErrors(6, i)
    cboSpecialist.Value = arrayErrors(7, i)
    cboProcess.Value = arrayErrors(8, i)
    cboError.Value = arrayErrors(9, i)
    cboCoach.Value = arrayErrors(10, i)
    'cboSubmitterId.Value = arrayErrors(11, i)
    txtCounter.Value = recordCounter + 1

End Sub
Private Sub btnForward_Click()
    If recordMax = 0 Then Exit Sub
    If recordCounter = recordMax Then
        MsgBox ("No more records")
        Exit Sub
    Else
        recordCounter = recordCounter + 1
        Call populateBoxes(recordCounter)
    End If
End Sub

Private Sub btnBack_Click()
    If recordMax = 0 Then Exit Sub
    If recordCounter = 0 Then
        MsgBox ("No more records")
        Exit Sub
    Else
        recordCounter = recordCounter - 1
        Call populateBoxes(recordCounter)
    End If
End Sub
Private Sub cboSpecialist_Change()
    Call ClearForm
End Sub

Private Sub cboWeekOf_Change()
    Call ClearForm
End Sub
Private Sub btnUpdate_Click()

    If ValidateForm(2) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If
    
On Error GoTo err1:
    If MsgBox("Do you want to update this record?", 48 + 4, "Submit Changes?") = vbYes Then
    
         '----UPDATE INFO HERE----
        Dim recordId As Long: recordId = txtID.Value
        Dim db As DAO.Database: Set db = CurrentDb
        Dim rs As DAO.Recordset
        Set rs = db.OpenRecordset("review_items_mkt", dbOpenDynaset)
            With rs
                .FindFirst ("ID = " & recordId)
                    If .NoMatch Then
                        MsgBox ("Record not located. Please search again")
                    Else
                        .Edit
                        ![member_id] = txtMemberID.Value
                        ![additional_id] = txtAdditionalId.Value
                        ![report_name] = txtReportName.Value
                        ![error_txt] = txtErrorComment.Value
                        ![submit_date] = Now()
                        ![week_id] = cboWeekOf.Value
                        ![employee_id] = cboSpecialist.Value
                        ![process_id] = cboProcess.Value
                        ![error_id] = cboError.Value
                        ![coach_id] = cboCoach.Value
                        ![submitter_id] = userId
                        .Update
                            arrayErrors(1, recordCounter) = txtMemberID.Value
                            arrayErrors(2, recordCounter) = txtAdditionalId.Value
                            arrayErrors(3, recordCounter) = txtReportName.Value
                            arrayErrors(4, recordCounter) = txtErrorComment.Value
                            arrayErrors(5, recordCounter) = txtSubmitDate.Value
                            arrayErrors(6, recordCounter) = cboWeekOf.Value
                            arrayErrors(7, recordCounter) = cboSpecialist.Value
                            arrayErrors(8, recordCounter) = cboProcess.Value
                            arrayErrors(9, recordCounter) = cboError.Value
                            arrayErrors(10, recordCounter) = cboCoach.Value
                            arrayErrors(11, recordCounter) = userId
                        MsgBox ("Item updated")
                    End If
            End With
        rs.Close: Set rs = Nothing
        db.Close: Set db = Nothing

        Call GetScore(cboSpecialist.Value, cboWeekOf.Value, cboProcess.Value)
        
    End If
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Record Locator; btnUpdate_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub
Private Sub btnDelete_Click()

    If ValidateForm(3) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If
    
On Error GoTo err1:
    If MsgBox("Are you sure you want to delete this record?", 48 + 4, "Delete record?") = vbYes Then
        
        '----DELETE INFO HERE----
        Dim recordId As Long: recordId = txtID.Value
        Dim db As DAO.Database: Set db = CurrentDb
        Dim rs As DAO.Recordset
        Set rs = db.OpenRecordset("review_items_mkt", dbOpenDynaset)
            With rs
                .FindFirst ("ID = " & recordId)
                    If .NoMatch Then
                        MsgBox ("Record not located. Please search again.")
                    Else
                        .Delete
                        MsgBox ("Item deleted")
                    End If
            End With
        rs.Close: Set rs = Nothing
        db.Close: Set db = Nothing
        Call GetScore(cboSpecialist.Value, cboWeekOf.Value, cboProcess.Value)
        Call btnSearch_Click
        
    End If
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Record Locator; btnDelete_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
End Sub
Private Sub btnClear_Click()
    Call ClearForm
End Sub

Private Sub ClearForm()
    txtID.Value = Null
    txtMemberID.Value = Null
    txtAdditionalId.Value = Null
    txtReportName.Value = Null
    txtErrorComment.Value = Null
    txtSubmitDate.Value = Null
    cboProcess.Value = Null
    cboError.Value = Null
    cboCoach.Value = Null
    txtCounter.Value = Null
    txtRecordCount.Value = Null
    txtErrorBox.Value = Null
    txtErrorBar.Value = Null
    txtErrorBar.BackColor = RGB(245, 246, 246)
End Sub

Private Sub Label43_Click()
    DoCmd.Close acForm, "Record Locator", acSaveNo
End Sub

Private Sub Label46_Click()
    DoCmd.OpenForm "Home", acNormal, , , acFormPropertySettings
End Sub

Public Function ValidateForm(submitType As Integer) As Boolean

    Dim msgStr As String, headerStr As String, footerStr As String, ctlName As String
    Dim varFocus As Variant
    
    headerStr = "<ul>"
    footerStr = "</ul>"
    
    Select Case submitType
        Case 1 '--search for records
            If cboWeekOf.Value = "" Or IsNull(cboWeekOf.Value) = True Then
                msgStr = msgStr & "<li><b>Week Of</b> cannot be blank.</li>"
                ctlName = "cboWeekOf,"
            End If
            If cboSpecialist.Value = "" Or IsNull(cboSpecialist.Value) = True Then
                msgStr = msgStr & "<li><b>Specialist</b> cannot be blank.</li>"
                ctlName = ctlName & "cboSpecialist,"
            End If
        Case 2 '--update record
            txtMemberID.Value = ScrubMemberId(txtMemberID.Value & vbNullString)
        
            If cboWeekOf.Value = "" Or IsNull(cboWeekOf.Value) = True Then
                msgStr = msgStr & "<li><b>Week Of</b> cannot be blank.</li>"
                ctlName = "cboWeekOf,"
            End If
            If cboSpecialist.Value = "" Or IsNull(cboSpecialist.Value) = True Then
                msgStr = msgStr & "<li><b>Specialist</b> cannot be blank.</li>"
                ctlName = ctlName & "cboSpecialist,"
            End If
            If cboProcess.Value = 100 Or IsNull(cboProcess.Value) = True Then
                msgStr = msgStr & "<li><b>Process</b> cannot be blank.</li>"
                ctlName = ctlName & "cboProcess,"
            End If
            If txtMemberID.Value = "" Or IsNull(txtMemberID.Value) = True Then
                msgStr = msgStr & "<li><b>Member ID</b> cannot be blank.</li>"
                ctlName = ctlName & "txtMemberId,"
            End If
            If txtReportName.Value = "" Or IsNull(txtReportName.Value) = True Then
                msgStr = msgStr & "<li><b>Report Name</b> cannot be blank.</li>"
                ctlName = ctlName & "txtReportName,"
            End If
            If cboError.Value = "" Or IsNull(cboError.Value) = True Then
                msgStr = msgStr & "<li><b>Error Category</b> cannot be blank.</li>"
                ctlName = ctlName & "cboError,"
            End If
            If (cboError.Value <> 0 And (IsNull(txtErrorComment.Value) Or txtErrorComment = "")) Then
                msgStr = msgStr & "<li><b>Comment</b> cannot be blank when submitting an error.</li>"
                ctlName = ctlName & "txtErrorComment,"
            End If
            If txtID.Value = "" Or IsNull(txtID) Then
                msgStr = msgStr & "<li><b>Update Record: </b> You must search for records before updating.</li>"
                ctlName = ctlName & "btnSearch,"
            End If
            If Len(txtMemberID.Value) <> 11 Then
                msgStr = msgStr & "<li><b>Member ID</b> must be 11 characters long.</li>"
                ctlName = ctlName & "txtMemberId,"
            End If
            If IsNumeric(txtMemberID.Value) <> True Then
                msgStr = msgStr & "<li><b>Member ID</b> can only be numbers. Please remove all other characters.</li>"
                ctlName = ctlName & "txtMemberId,"
            End If
        Case 3 '--delete record
            If cboWeekOf.Value = "" Or IsNull(cboWeekOf.Value) = True Then
                msgStr = msgStr & "<li><b>Week Of</b> cannot be blank.</li>"
                ctlName = "cboWeekOf,"
            End If
            If cboSpecialist.Value = "" Or IsNull(cboSpecialist.Value) = True Then
                msgStr = msgStr & "<li><b>Specialist</b> cannot be blank.</li>"
                ctlName = ctlName & "cboSpecialist,"
            End If
            If txtID.Value = "" Or IsNull(txtID) Then
                msgStr = msgStr & "<li><b>Delete Record: </b> You must search for records before deleting.</li>"
                ctlName = ctlName & "btnSearch,"
            End If
    End Select
    
    If msgStr = "" Then 'no errors
        txtErrorBox.Value = Null
        txtErrorBar.Value = Null
        txtErrorBar.BackColor = RGB(245, 246, 246)
        ValidateForm = False
    Else 'errors, show msg to user
        txtErrorBox.Value = headerStr & msgStr & footerStr
        txtErrorBar.Value = "Submission Errors"
        txtErrorBar.BackColor = RGB(255, 186, 0)
        varFocus = Split(ctlName, ",")
        Me.Controls(varFocus(0)).SetFocus
        ValidateForm = True
    End If

End Function
