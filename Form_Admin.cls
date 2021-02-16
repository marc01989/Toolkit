VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnSubmitExternalQA_Click()
    If ValidateForm(3) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If
    
On Error GoTo err1:
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim currScore As Double, newScore As Double
    Set db = CurrentDb
    Set rs = db.OpenRecordset("external_qa")
        With rs
            .AddNew
                ![week_id] = cboWeekOf.Value
                ![employee_id] = cboSpecialist.Value
                ![audit_type] = cboAuditType.Value
                ![accuracy] = txtAccuracy.Value
                ![audit_date] = txtAuditDate.Value
            .Update
        End With
    rs.Close

    Set rs = db.OpenRecordset("scores_mkt")
    With rs
        .FindFirst ("week_id = " & cboWeekOf.Value & " AND employee_id = " & cboSpecialist.Value)
        If .NoMatch Then
            .AddNew
                ![week_id] = cboWeekOf.Value
                ![employee_id] = cboSpecialist.Value
                ![external_qa] = txtAccuracy.Value
                ![submit_date] = Now()
            .Update
        Else
            .Edit
                If Not IsNull(![external_qa]) Then
                    newScore = Round((CDbl(![external_qa]) + CDbl(txtAccuracy.Value)) / 2, 2)
                    ![external_qa] = newScore
                Else
                    ![external_qa] = txtAccuracy.Value
                End If
            .Update
        End If
    End With

    MsgBox ("Item Submitted")
    txtAccuracy.Value = Null
    txtAuditDate.Value = Null
    
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Admin; btnSubmitExternalQA_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select

End Sub

Private Sub Form_Load()
    cboCoach.RowSource = "SELECT employee_id, employee_name FROM employees WHERE is_coach = 1 ORDER BY employee_name UNION SELECT 0, 'Select Coach' FROM employees;"
    cboCoach.Value = 0
    cboIsNewHire.RowSource = "SELECT status_id, status FROM status UNION SELECT 2, 'New Hire?' FROM status;"
    cboIsNewHire.Value = 2
    cboIsCoach.RowSource = "SELECT status_id, status FROM status UNION SELECT 2, 'Is Leadership?' FROM status;"
    cboIsCoach.Value = 2
End Sub

Private Sub btnImport_Click()
    If cboWeekOf.Value = "" Or IsNull(cboWeekOf) Then
        If MsgBox("You must select a 'Week Of'.", vbExclamation + vbOKOnly, "Cannot Import") = vbOK Then
            Exit Sub
        End If
    End If

    Select Case optImportType
        Case 1
            Call ImportProduction
        Case 2
            Call ImportAuxStatus
    End Select
End Sub

Private Sub btnSubmitAnnouncement_Click()
    MsgBox ("doesn't work")
End Sub

Private Sub ImportProduction()
    DoCmd.SetWarnings False
    If DCount("[week_id]", "[weekly_production]", "[week_id] = " & cboWeekOf.Value) = 0 Then
        If MsgBox("Import Production for the week of " & cboWeekOf.Column(1) & "?", vbYesNo, "Import Production") = vbYes Then
            DoCmd.runSQL ("INSERT INTO weekly_production ( week_id, employee_id, segment_id, amount ) " & _
            " SELECT " & cboWeekOf.Value & ", employees.employee_id, production_segments.segment_id, lnkWeeklyProduction.Value " & _
            " FROM production_segments INNER JOIN (lnkWeeklyProduction INNER JOIN employees ON lnkWeeklyProduction.Row = employees.employee_name) ON production_segments.segment = lnkWeeklyProduction.Column;")
            MsgBox ("Job Complete")
        Else
            Exit Sub
        End If
    Else
        If MsgBox("Production records exist for the week of " & cboWeekOf.Column(1) & ". Do you want to overwrite them?", vbYesNo) = vbYes Then
            DoCmd.runSQL ("DELETE * FROM weekly_production WHERE week_id = " & cboWeekOf.Value & ";")
            DoCmd.runSQL ("INSERT INTO weekly_production ( week_id, employee_id, segment_id, amount ) " & _
            " SELECT " & cboWeekOf.Value & ", employees.employee_id, production_segments.segment_id, lnkWeeklyProduction.Value " & _
            " FROM production_segments INNER JOIN (lnkWeeklyProduction INNER JOIN employees ON lnkWeeklyProduction.Row = employees.employee_name) ON production_segments.segment = lnkWeeklyProduction.Column;")
            MsgBox ("Job Complete")
        Else
            Exit Sub
        End If
    End If
End Sub

Private Sub ImportAuxStatus()
    If MsgBox("Import Auxiliary Status Times for the week of " & cboWeekOf.Column(1) & "?", vbYesNo, "Import Aux") = vbYes Then
        DoCmd.runSQL ("INSERT INTO auxiliary_status_data ( week_id, employee_id, media_id, Cal_Date2, reason_id, textbox3 ) " & _
        " SELECT " & cboWeekOf.Value & ", employees.employee_id, media.media_id, lnkAuxiliaryStatuses.Cal_Date2, auxiliary_statuses.reason_id, lnkAuxiliaryStatuses.textbox3 " & _
        " FROM auxiliary_statuses INNER JOIN (media INNER JOIN (employees INNER JOIN lnkAuxiliaryStatuses ON employees.network_id_full = lnkAuxiliaryStatuses.AGENT_NAME3) ON media.MEDIA_NAME2 = lnkAuxiliaryStatuses.MEDIA_NAME2) ON auxiliary_statuses.REASON2 = lnkAuxiliaryStatuses.REASON2;")
        MsgBox ("Job Complete")
    End If
End Sub
Private Sub Label12_Click()
    DoCmd.Close acForm, "Admin", acSaveNo
End Sub

Private Sub Label46_Click()
    DoCmd.OpenForm "Home", acNormal, , , acFormPropertySettings
End Sub
Private Sub cboSpecialist1_AfterUpdate()

On Error GoTo err1:
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT employees.*, employees_1.employee_id " & _
    " FROM employees AS employees_1 INNER JOIN employees ON employees_1.employee_id = employees.coach_id " & _
    " WHERE employees.employee_id = " & cboSpecialist1.Value & ";")
        With rs
            If .recordCount > 0 Then
                txtEmployeeId1.Value = ![employees.employee_id]
                cboCoach1.Value = ![employees_1.employee_id]
                cboIsActive1.Value = ![is_active]
                txtNetworkId1.Value = ![network_id]
                txtHireDate1.Value = ![hire_date]
                cboIsNewHire1.Value = ![is_new_hire]
                txtTitle1.Value = ![Title]
                cboIsCoach1.Value = ![is_coach]
                txtLastPromotionDate1.Value = ![last_promotion_date]
            End If
        End With
        rs.Close: Set rs = Nothing
        db.Close: Set db = Nothing
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "QA Input; btnSubmit_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
End Sub
Private Sub btnClear_Click()
    Call ClearForm(1)
End Sub
'UPDATE EMPLOYEE
Private Sub btnUpdate_Click()

    If ValidateForm(1) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If
    
On Error GoTo err1:
    If MsgBox("Would you like to submit the changes?", vbYesNoCancel, "Submit?") = vbYes Then
        Dim db As DAO.Database: Set db = CurrentDb
        Dim rs As DAO.Recordset
        Dim empName As Variant
        Dim netIDFull As String
        Set rs = db.OpenRecordset("employees")
            With rs
                .FindFirst ("employee_id = " & txtEmployeeId1.Value)
                    If InStr(1, cboSpecialist1.Column(1), " ") <> 0 Then
                        empName = Split(cboSpecialist1.Column(1), " ")
                        netIDFull = empName(1) & ", " & empName(0) & " (" & txtNetworkId1.Value & ")"
                    End If
                    If Not .NoMatch Then
                        .Edit
                        ![coach_id] = cboCoach1.Value
                        ![is_active] = cboIsActive1.Value
                        ![network_id_full] = netIDFull
                        ![network_id] = txtNetworkId1.Value
                        ![hire_date] = txtHireDate1.Value
                        ![is_new_hire] = cboIsNewHire1.Value
                        ![Title] = txtTitle1.Value
                        ![is_coach] = cboIsCoach1.Value
                        ![last_promotion_date] = txtLastPromotionDate1.Value
                        ![user_name] = txtNetworkId1.Value
                        .Update
                    Else
                        MsgBox ("Specialist not found")
                    End If
                    rs.Close: Set rs = Nothing
                    db.Close: Set db = Nothing
            End With
        MsgBox ("Item submitted.")
        Call ClearForm(1)
    End If
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Admin; btnUpdate_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
End Sub
'ADD NEW
Private Sub btnAddNewEmployee_Click()

    If ValidateForm(2) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If
    
On Error GoTo err1:
    If MsgBox("Would you like to add a new employee?", vbYesNoCancel, "Submit?") = vbYes Then
        Dim empName As Variant
        Dim netIDFull As String
        Dim db As DAO.Database: Set db = CurrentDb
        Dim rs As DAO.Recordset
        Set rs = db.OpenRecordset("employees")
            If InStr(1, txtName.Value, " ") <> 0 Then
                empName = Split(txtName, " ")
                netIDFull = empName(1) & ", " & empName(0) & " (" & txtNetworkId.Value & ")"
            End If
            With rs
                .AddNew
                    ![employee_id] = DMax("employee_id", "employees", "employee_id > 0") + 1
                    ![employee_name] = txtName.Value
                    ![coach_id] = cboCoach.Value
                    ![is_active] = 1
                    ![network_id_full] = netIDFull
                    ![network_id] = txtNetworkId.Value
                    ![hire_date] = txtHireDate.Value
                    ![is_new_hire] = cboIsNewHire.Value
                    ![Title] = txtTitle.Value
                    ![is_coach] = cboIsCoach.Value
                    ![is_manager] = 0
                    ![last_promotion_date] = txtLastPromotionDate.Value
                    ![user_name] = txtNetworkId.Value
                .Update
                rs.Close: Set rs = Nothing
                db.Close: Set db = Nothing
            End With
        MsgBox ("Item submitted.")
        Call ClearForm(2)
    End If
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Admin; btnAddNewEmployee_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
End Sub

Public Sub ClearForm(clearType As Integer)
    Select Case clearType
    Case 1 '--UPDATE EXISTING--
        cboCoach1.Value = Null
        cboIsActive1.Value = Null
        txtNetworkId1.Value = Null
        txtHireDate1.Value = Null
        cboIsNewHire1.Value = Null
        txtTitle1.Value = Null
        cboIsCoach1.Value = Null
        txtLastPromotionDate1.Value = Null
        txtEmployeeId1.Value = Null
        cboSpecialist1.Value = Null
    Case 2 '--ADD NEW--
        txtName.ForeColor = RGB(166, 166, 166)
        txtName.Value = Null
        cboCoach.RowSource = "SELECT employee_id, employee_name FROM employees WHERE is_coach = 1 ORDER BY employee_name UNION SELECT 0, 'Select Coach' FROM employees;"
        cboCoach.Value = 0
        txtTitle.Value = Null
        txtTitle.ForeColor = RGB(166, 166, 166)
        cboIsNewHire.RowSource = "SELECT status_id, status FROM status UNION SELECT 2, 'New Hire?' FROM status;"
        cboIsNewHire.Value = 2
        cboIsNewHire.ForeColor = RGB(166, 166, 166)
        txtNetworkId.Value = Null
        txtNetworkId.ForeColor = RGB(166, 166, 166)
        cboIsCoach.RowSource = "SELECT status_id, status FROM status UNION SELECT 2, 'Is Specialist Leadership?' FROM status;"
        cboIsCoach.Value = 2
        cboIsCoach.ForeColor = RGB(166, 166, 166)
        txtHireDate.Value = Null
        txtHireDate.ForeColor = RGB(166, 166, 166)
        txtLastPromotionDate.ForeColor = RGB(166, 166, 166)
        txtLastPromotionDate.Value = Null
    End Select
    txtErrorBar.Value = Null
    txtErrorBar.BackColor = RGB(245, 246, 246)
    txtErrorBox.Value = Null
End Sub
Private Sub txtName_Exit(Cancel As Integer)
    If txtName.Value = "" Or IsNull(txtName) Then
        txtName.ForeColor = RGB(166, 166, 166)
    End If
End Sub
Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    txtName.ForeColor = vbBlack
End Sub

Private Sub txtTitle_Exit(Cancel As Integer)
    If txtTitle.Value = "" Or IsNull(txtTitle) Then
        txtTitle.ForeColor = RGB(166, 166, 166)
    End If
End Sub
Private Sub txtTitle_KeyDown(KeyCode As Integer, Shift As Integer)
    txtTitle.ForeColor = vbBlack
End Sub
Private Sub txtNetworkId_Exit(Cancel As Integer)
    If txtNetworkId.Value = "" Or IsNull(txtNetworkId) Then
        txtNetworkId.ForeColor = RGB(166, 166, 166)
    End If
End Sub
Private Sub txtNetworkId_KeyDown(KeyCode As Integer, Shift As Integer)
    txtNetworkId.ForeColor = vbBlack
End Sub
Private Sub txtHireDate_Exit(Cancel As Integer)
    If txtHireDate.Value = "" Or IsNull(txtHireDate) Then
        txtHireDate.ForeColor = RGB(166, 166, 166)
    End If
End Sub
Private Sub txtHireDate_KeyDown(KeyCode As Integer, Shift As Integer)
    txtHireDate.ForeColor = vbBlack
End Sub
Private Sub txtLastPromotionDate_Exit(Cancel As Integer)
    If txtLastPromotionDate.Value = "" Or IsNull(txtLastPromotionDate) Then
        txtLastPromotionDate.ForeColor = RGB(166, 166, 166)
    End If
End Sub
Private Sub txtLastPromotionDate_KeyDown(KeyCode As Integer, Shift As Integer)
    txtLastPromotionDate.ForeColor = vbBlack
End Sub



Private Sub cboCoach_Enter()
    cboCoach.RowSource = "SELECT employee_id, employee_name FROM employees WHERE is_coach = 1 ORDER BY employee_name;"
End Sub
Private Sub cboCoach_Exit(Cancel As Integer)
    If IsNull(cboCoach) Or cboCoach.Value = 0 Then
        cboCoach.RowSource = "SELECT employee_id, employee_name FROM employees WHERE is_coach = 1 ORDER BY employee_name UNION SELECT 0, 'Select Coach' FROM employees;"
        cboCoach.Value = 0
    End If
End Sub
Private Sub cboIsNewHire_Enter()
    cboIsNewHire.RowSource = "SELECT status_id, status FROM status;"
    cboIsNewHire.ForeColor = vbBlack
End Sub
Private Sub cboIsNewHire_Exit(Cancel As Integer)
    If IsNull(cboIsNewHire) Or cboIsNewHire.Value = 2 Then
        cboIsNewHire.RowSource = "SELECT status_id, status FROM status UNION SELECT 2, 'New Hire?' FROM status;"
        cboIsNewHire.Value = 2
        cboIsNewHire.ForeColor = RGB(166, 166, 166)
    End If
End Sub
Private Sub cboIsCoach_Enter()
    cboIsCoach.RowSource = "SELECT status_id, status FROM status;"
    cboIsCoach.ForeColor = vbBlack
End Sub
Private Sub cboIsCoach_Exit(Cancel As Integer)
    If IsNull(cboIsCoach) Or cboIsCoach.Value = 2 Then
        cboIsCoach.RowSource = "SELECT status_id, status FROM status UNION SELECT 2, 'Is Leadership?' FROM status;"
        cboIsCoach.Value = 2
        cboIsCoach.ForeColor = RGB(166, 166, 166)
    End If
End Sub

Public Function ValidateForm(submitType As Integer) As Boolean
'REFERENCE REQUIRED - Microsoft VBScript Regular Expression 5.5

    Dim msgStr As String, headerStr As String, footerStr As String, ctlName As String
    Dim varFocus As Variant
    Dim regex As RegExp: Set regex = New RegExp
    headerStr = "<ul>": footerStr = "</ul>"
    With regex
        .MultiLine = False
        .Global = True
        .IgnoreCase = True
        .Pattern = "^(0[1-9]|1[012])[\/](0[1-9]|[12][0-9]|3[01])[\/](19|20)\d\d$"
    End With

    Select Case submitType
        Case 1 '--UPDATE EXISTING--
            If txtTitle1.Value = "" Or IsNull(txtTitle1.Value) = True Then
                msgStr = msgStr & "<li><b>Title</b> cannot be blank.</li>"
                ctlName = ctlName & "txtTitle1,"
            End If
            If cboCoach1.Value = "" Or IsNull(cboCoach1.Value) = True Then
                msgStr = msgStr & "<li><b>Coach Name</b> cannot be blank.</li>"
                ctlName = ctlName & "cboCoach1,"
            End If
            If cboIsActive1.Value = "" Or IsNull(cboIsActive1.Value) = True Then
                msgStr = msgStr & "<li><b>Is Active</b> cannot be blank.</li>"
                ctlName = ctlName & "cboIsActive1,"
            End If
            If cboIsCoach1.Value = "" Or IsNull(cboIsCoach1.Value) = True Then
                msgStr = msgStr & "<li><b>Is Specialist a Coach?</b> cannot be blank.</li>"
                ctlName = ctlName & "cboIsCoach1,"
            End If
            If txtNetworkId1.Value = "" Or IsNull(txtNetworkId1.Value) = True Then
                msgStr = msgStr & "<li><b>Network Id</b> cannot be blank.</li>"
                ctlName = ctlName & "txtNetworkId1,"
            End If
            If Not IsNull(txtHireDate1) Then
                If regex.test(txtHireDate1.Value) = False Then
                    msgStr = msgStr & "<li><b>Hire Date</b> must be mm/dd/yyyy format.</li>"
                    ctlName = ctlName & "txtHireDate1,"
                End If
            End If
            If Not IsNull(txtLastPromotionDate1) Then
                If regex.test(txtLastPromotionDate1.Value) = False Then
                    msgStr = msgStr & "<li><b>Last Promotion Date</b> must be mm/dd/yyyy format.</li>"
                    ctlName = ctlName & "txtLastPromotionDate1,"
                End If
            End If
        Case 2 '--ADD NEW--
            If txtName.Value = "" Or IsNull(txtName.Value) = True Then
                msgStr = msgStr & "<li><b>Employee Name</b> cannot be blank.</li>"
                ctlName = "txtName,"
            End If
            If cboCoach.Value = 0 Or IsNull(cboCoach.Value) = True Then
                msgStr = msgStr & "<li><b>Coach Name</b> cannot be blank.</li>"
                ctlName = ctlName & "cboCoach,"
            End If
            If txtTitle.Value = "" Or IsNull(txtTitle.Value) = True Then
                msgStr = msgStr & "<li><b>Title</b> cannot be blank.</li>"
                ctlName = ctlName & "txtTitle,"
            End If
            If cboIsNewHire.Value = 2 Or IsNull(cboIsNewHire) Then
                msgStr = msgStr & "<li><b>New Hire?</b> cannot be blank .</li>"
                ctlName = ctlName & "cboIsNewHire,"
            End If
            If txtNetworkId.Value = "" Or IsNull(txtNetworkId.Value) = True Then
                msgStr = msgStr & "<li><b>Network Id</b> cannot be blank.</li>"
                ctlName = ctlName & "txtNetworkId,"
            End If
            If cboIsCoach.Value = 2 Or IsNull(cboIsCoach.Value) = True Then
                msgStr = msgStr & "<li><b>Is Leadership?</b> cannot be blank.</li>"
                ctlName = ctlName & "cboIsCoach,"
            End If
            If Not IsNull(txtHireDate) Then
                If regex.test(txtHireDate.Value) = False Then
                    If msgStr = "" Then: msgStr = "<li><b>Hire Date</b> must be mm/dd/yyyy format.</li>"
                    ctlName = ctlName & "txtHireDate,"
                End If
            End If
            If Not IsNull(txtLastPromotionDate1) Then
                If regex.test(txtLastPromotionDate1.Value) = False Then
                    If msgStr = "" Then: msgStr = msgStr & "<li><b>Last Promotion Date</b> must be mm/dd/yyyy format.</li>"
                    ctlName = ctlName & "txtLastPromotionDate1,"
                End If
            End If
        Case 3 'EXTERNAL QA
            If cboWeekOf.Value = "" Or IsNull(cboWeekOf.Value) = True Then
                msgStr = msgStr & "<li><b>Week Of</b> cannot be blank.</li>"
                ctlName = ctlName & "cboWeekOf,"
            End If
            If cboSpecialist.Value = "" Or IsNull(cboSpecialist.Value) = True Then
                msgStr = msgStr & "<li><b>Specialist</b> cannot be blank.</li>"
                ctlName = ctlName & "cboSpecialist,"
            End If
            If txtAccuracy.Value = "" Or IsNull(txtAccuracy.Value) = True Then
                msgStr = msgStr & "<li><b>Accuracy</b> cannot be blank.</li>"
                ctlName = ctlName & "txtAccuracy,"
            End If
            If cboAuditType.Value = "" Or IsNull(cboAuditType.Value) = True Then
                msgStr = msgStr & "<li><b>Audit Type</b> cannot be blank.</li>"
                ctlName = ctlName & "cboAuditType,"
            End If
            If txtAuditDate.Value = "" Or IsNull(txtAuditDate.Value) = True Then
                msgStr = msgStr & "<li><b>Audit Date</b> cannot be blank.</li>"
                ctlName = ctlName & "txtAuditDate,"
            End If
            regex.Pattern = ("^(?:\d{0,3}\.\d{1,2})$|^\d{1,3}$") 'match decimal format
            If Not IsNull(txtAccuracy.Value) Then
                If regex.test(txtAccuracy.Value) = False Then
                    msgStr = msgStr & "<li><b>Accuracy </b> must be a decimal number.</li>"
                    ctlName = ctlName & "txtAccuracy,"
                End If
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


