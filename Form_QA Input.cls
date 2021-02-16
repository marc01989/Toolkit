VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_QA Input"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
    Dim userId As Integer

Private Sub cboSpecialist_AfterUpdate()
    lstReviewLog.Requery
End Sub

Private Sub cboWeekOf_AfterUpdate()
    lstReviewLog.Requery
End Sub

Private Sub Form_Load()
    Dim obj As AccessObject, dbs As Object
    Dim args As String
    
    'checks if value is passed when opening form, runs main query if so
    If Not IsNull(Forms![QA Input].OpenArgs) Then
        args = Forms![QA Input].OpenArgs
            If args = 0 Or IsNull(args) Then
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
                userId = CInt(args)
                'MsgBox (userId)
                Call btnClearForm_Click
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


Private Sub btnClearForm_Click()
    ClearForm (1)
End Sub
Private Sub btnOpenRecordLocator_Click()
    If CurrentProject.AllForms("Record Locator").IsLoaded = False Then
        DoCmd.OpenForm "Record Locator", acNormal, , , , acWindowNormal, Trim(Str(cboWeekOf.Value) & "," & Str(cboSpecialist.Value) & ", " & userId)
    Else
        DoCmd.Close acForm, "Record Locator"
        DoCmd.OpenForm "Record Locator", acNormal, , , , acWindowNormal, Trim(Str(cboWeekOf.Value) & "," & Str(cboSpecialist.Value) & ", " & userId)
    End If
End Sub

Private Sub btnSubmit_Click()

    If ValidateForm(1) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If
    
On Error GoTo err1:

    '----SUBMIT INFO HERE----
    Dim errorId As Integer
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("review_items_mkt", dbOpenDynaset)
    If optFYI.Value = 1 Then: errorId = cboError.Value: Else: errorId = 30
    
        With rs
            .AddNew
                ![member_id] = txtMemberID.Value
                ![additional_id] = txtAdditionalId.Value
                ![report_name] = txtReportName.Value
                ![error_txt] = txtErrorComment.Value
                ![submit_date] = Now()
                ![week_id] = cboWeekOf.Value
                ![employee_id] = cboSpecialist.Value
                ![process_id] = cboProcess.Value
                ![error_id] = errorId
                ![coach_id] = DLookup("[coach_id]", "[employees]", "[employee_id] = " & cboSpecialist.Value)
                ![submitter_id] = userId
            .Update
        End With
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    
    Call GetScore(cboSpecialist.Value, cboWeekOf.Value, cboProcess.Value)
    
    MsgBox ("Item submitted")
    ClearForm (2)
    lstReviewLog.Requery
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "QA Input; btnSubmit_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select

End Sub
Private Sub Label46_Click()
    DoCmd.OpenForm "Home", acNormal, , , acFormPropertySettings
End Sub

Private Sub Label60_Click()
    DoCmd.Close acForm, "QA Input", acSaveNo
End Sub
Private Sub txtMemberId_Exit(Cancel As Integer)
    If txtMemberID.Value = "" Or IsNull(txtMemberID) Then
        txtMemberID.ForeColor = RGB(166, 166, 166)
    End If
End Sub
Private Sub txtMemberId_KeyDown(KeyCode As Integer, Shift As Integer)
    txtMemberID.ForeColor = vbBlack
End Sub
Private Sub txtReportName_Exit(Cancel As Integer)
    If txtReportName.Value = "" Or IsNull(txtReportName) Then
        txtReportName.ForeColor = RGB(166, 166, 166)
    End If
End Sub
Private Sub txtReportName_KeyDown(KeyCode As Integer, Shift As Integer)
    txtReportName.ForeColor = vbBlack
End Sub
Private Sub txtAdditionalId_Exit(Cancel As Integer)
    If txtAdditionalId.Value = "" Or IsNull(txtAdditionalId) Then
        txtAdditionalId.ForeColor = RGB(166, 166, 166)
    End If
End Sub
Private Sub txtAdditionalId_KeyDown(KeyCode As Integer, Shift As Integer)
    txtAdditionalId.ForeColor = vbBlack
End Sub
Private Sub cboProcess_Enter()
    cboProcess.RowSource = "SELECT process_id, [processes].[process_name] FROM processes;"
    cboProcess.ForeColor = vbBlack
End Sub

Private Sub cboProcess_Exit(Cancel As Integer)
    If IsNull(cboProcess) Or cboProcess.Value = 100 Then
        cboProcess.RowSource = "SELECT process_id, [processes].[process_name] FROM processes UNION SELECT 100, 'Select One' FROM processes;"
        cboProcess.Value = 100
        cboProcess.ForeColor = RGB(166, 166, 166)
    End If
End Sub
Private Sub ClearForm(clearType As Integer)

    Select Case clearType
        Case 1 '--clear button
            cboError.Value = 0
            cboWeekOf.Value = Null
            cboSpecialist.Value = Null
            cboProcess.Value = Null
            txtMemberID.Value = Null
            txtMemberID.ForeColor = RGB(166, 166, 166)
            txtReportName.Value = Null
            txtReportName.ForeColor = RGB(166, 166, 166)
            txtAdditionalId.Value = Null
            txtAdditionalId.ForeColor = RGB(166, 166, 166)
            txtErrorComment.Value = Null
            txtErrorBar.Value = Null
            txtErrorBar.BackColor = RGB(245, 246, 246)
            txtErrorBox.Value = Null
            optFYI.Value = 1
            lstReviewLog.Requery
            cboProcess.RowSource = "SELECT process_id, [processes].[process_name] FROM processes UNION SELECT 100, 'Select One' FROM processes;"
            cboProcess.Value = 100
            cboProcess.ForeColor = RGB(166, 166, 166)
        Case 2 '--on submit
            cboError.Value = 0
            txtMemberID.Value = Null
            txtMemberID.ForeColor = RGB(166, 166, 166)
            txtReportName.Value = Null
            txtReportName.ForeColor = RGB(166, 166, 166)
            txtAdditionalId.Value = Null
            txtAdditionalId.ForeColor = RGB(166, 166, 166)
            txtErrorComment.Value = Null
            txtErrorBar.Value = Null
            txtErrorBar.BackColor = RGB(245, 246, 246)
            txtErrorBox.Value = Null
            optFYI.Value = 1
            lstReviewLog.Requery
    End Select
End Sub

Public Function ValidateForm(submitType As Integer) As Boolean

    Dim msgStr As String, headerStr As String, footerStr As String, ctlName As String
    Dim varFocus As Variant
    
    headerStr = "<ul>"
    footerStr = "</ul>"
    
    Select Case submitType
        Case 1
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
            If Len(txtMemberID.Value) <> 11 Then
                msgStr = msgStr & "<li><b>Member ID</b> must be 11 characters long.</li>"
                ctlName = ctlName & "txtMemberId,"
            End If
            If IsNumeric(txtMemberID.Value) <> True Then
                msgStr = msgStr & "<li><b>Member ID</b> can only be numbers. Please remove all other characters.</li>"
                ctlName = ctlName & "txtMemberId,"
            End If
        Case 2
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



