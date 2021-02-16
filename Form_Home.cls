VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim userId As Integer


Private Sub Form_Load()
    Dim args As String
    
    'checks if value is passed when opening form, runs main query if so
    If Not IsNull(Forms![Home].OpenArgs) Then
        args = Forms![Home].OpenArgs
            If args = "" Or IsNull(args) Then
                MsgBox ("User not found - please login again")
                DoCmd.OpenForm "Login", acNormal, , , , acWindowNormal
                DoCmd.Close acForm, "Home"
                Call LogError(0 & " " & "User Id not found or not passed to home screen", "Home; Form_Load()")
                Exit Sub
            Else
                userId = CInt(args)
                'MsgBox (userId)
            End If
    Else
        MsgBox ("User not found - please login again")
        DoCmd.OpenForm "Login", acNormal, , , , acWindowNormal
        DoCmd.Close acForm, "Home"
        Call LogError(0 & " " & "User Id not found or not passed to home screen", "Home; Form_Load()")
        Exit Sub
    End If


'--UPDATE COACHING LOG--
    Dim weekId As Integer: weekId = getCurrentWeekId()
    DoCmd.SetWarnings (False)

    If IsNull(DSum("[week_id]", "[coaching_log]", "[week_id] = " & weekId)) Then
        DoCmd.runSQL "INSERT INTO coaching_log (employee_id, coach_id) " & _
            " SELECT employee_id, coach_id FROM employees " & _
            " WHERE (is_manager = 0 AND is_active = 1);"
        DoCmd.runSQL "UPDATE coaching_log SET week_id = " & weekId & " WHERE week_id IS NULL;"
    End If
    
End Sub
Private Sub cboSpecialist_AfterUpdate()
    ClearForm
    If Not IsNull(cboSpecialist) And Not IsNull(cboWeekOf) Then
        Call GetNotes(cboWeekOf, cboSpecialist)
    End If
End Sub
Private Sub cboWeekOf_AfterUpdate()
    ClearForm
    If Not IsNull(cboSpecialist) And Not IsNull(cboWeekOf) Then
        Call GetNotes(cboWeekOf, cboSpecialist)
    End If
End Sub
Public Sub GetNotes(weekId As Integer, employeeId As Integer)
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT area_of_excellence, employee_comments " & _
    " FROM coaching_log WHERE [week_id]= " & weekId & " AND [employee_id]= " & employeeId & ";")
    With rs
        If .recordCount > 0 Then
            txtAreaOfExcellence.Value = ![area_of_excellence]
            txtEmployeeComments.Value = ![employee_comments]
        End If
    End With
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    txtAreaOfExcellence_Exit (0)
    
    
End Sub
Private Sub btnPrint_Click()

    If ValidateForm(1) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If
    
On Error GoTo err1:
    
    '----SUBMIT INFO TO scorecard_prep TABLE----
    DoCmd.SetWarnings False
    DoCmd.runSQL ("DELETE * FROM scorecard_prep;")
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim varWkScores(8) As Variant
    Dim var6Month As Variant: var6Month = Get6MonthScore(cboSpecialist.Value, cboWeekOf.Value)
    Dim coachId As Integer: coachId = DLookup("[coach_id]", "[employees]", "[employee_id] = " & cboSpecialist.Value)
    Dim coachName As String: coachName = DLookup("[employee_name]", "[employees]", "[employee_id] = " & coachId)
    Dim boolWkScores As Boolean
    Dim specialistname As String: specialistname = cboSpecialist.Column(1)

    'grab weekly scores, if they exist, and put into array
    Set db = CurrentDb
    Set rs = db.OpenRecordset("scores_mkt", dbOpenDynaset)
    With rs
        .FindFirst ("week_id = " & cboWeekOf.Value & " AND employee_id = " & cboSpecialist.Value)
        If .NoMatch Then
            boolWkScores = False
        Else
            boolWkScores = True
            'week avg
            varWkScores(0) = rs![hics]
            varWkScores(1) = rs![on_term_job]
            varWkScores(2) = rs![off_term_job]
            varWkScores(3) = rs![cancel_job]
            varWkScores(4) = rs![other]
            varWkScores(5) = rs![recon]
            varWkScores(6) = rs![cutlog_demo_changes]
            varWkScores(7) = rs![chat]
            varWkScores(8) = rs![external_qa]
        End If
    End With
    rs.Close

    'begin appending data to scorecard_prep table
    Set rs = db.OpenRecordset("scorecard_prep", dbOpenDynaset)
    With rs
        .AddNew
            'basic weekly info
            rs![week_start] = cboWeekOf.Column(1)
            rs![week_id] = cboWeekOf.Value
            rs![employee_name] = cboSpecialist.Column(1)
            rs![employee_id] = cboSpecialist.Value
            rs![coach_name] = DLookup("[employee_name]", "[employees]", "[employee_id] = " & coachId)
            rs![area_of_excellence] = txtAreaOfExcellence.Value
            rs![coaching_concerns] = DLookup("employee_comments", "coaching_log", "week_id = " & cboWeekOf.Value & " AND employee_id = " & cboSpecialist.Value)
            
            'if weekly scores exist, append to table
            If boolWkScores = True Then
            rs![hics] = varWkScores(0)
            rs![on_term_job] = varWkScores(1)
            rs![off_term_job] = varWkScores(2)
            rs![cancel_job] = varWkScores(3)
            rs![other_reporting] = varWkScores(4)
            rs![recon] = varWkScores(5)
            rs![cutlog_demo_changes] = varWkScores(6)
            rs![chat] = varWkScores(7)
            rs![external_qa] = varWkScores(8)
            End If
            
            '6 month avg
            rs![hics_avg] = var6Month(0)
            rs![on_term_job_avg] = var6Month(1)
            rs![off_term_job_avg] = var6Month(2)
            rs![cancel_job_avg] = var6Month(3)
            rs![other_reporting_avg] = var6Month(4)
            rs![recon_avg] = var6Month(5)
            rs![cutlog_demo_changes_avg] = var6Month(6)
            rs![chat_avg] = var6Month(7)
            'rs![external_qa_avg] = var6Month(8)
        .Update
    End With
    rs.Close

    'update coaching log with txtAreaOfExcellence.Value
    Set rs = db.OpenRecordset("coaching_log")
        With rs
            .FindFirst "week_id = " & cboWeekOf.Value & " AND employee_id = " & cboSpecialist.Value
            If Not .NoMatch Then
                .Edit
                    ![area_of_excellence] = txtAreaOfExcellence.Value
                .Update
            End If
        End With
    rs.Close: Set rs = Nothing
        
    'REFRESH qryScorecardAuxiliaryStatus QUERY, WHICH IS CONTROL SOURCE OF 'AUX STATUS' REPORT
    Dim qd As DAO.QueryDef
    Set qd = db.QueryDefs("qryScorecardAuxiliaryStatus")
    sql = "TRANSFORM (24*Int(Sum(textbox3))+Format(Sum(textbox3),'h')) & Format(Sum(textbox3),':nn:ss') AS [Total Not Ready]" & _
    " SELECT media.MEDIA_NAME2 AS Channel, auxiliary_status_data.Cal_Date2 AS [Date] " & _
    " FROM ((auxiliary_status_data INNER JOIN employees ON auxiliary_status_data.employee_id = employees.employee_id) INNER JOIN media ON auxiliary_status_data.media_id = media.media_id) INNER JOIN auxiliary_statuses ON auxiliary_status_data.reason_id = auxiliary_statuses.reason_ID " & _
    " WHERE ((([auxiliary_status_data.employee_id]) = " & cboSpecialist.Value & ") And ((auxiliary_status_data.textbox3) <> 0) And (([auxiliary_status_data.week_id]) = " & cboWeekOf.Value & ")) " & _
    " GROUP BY media.MEDIA_NAME2, auxiliary_status_data.Cal_Date2 " & _
    " ORDER BY media.MEDIA_NAME2, auxiliary_status_data.Cal_Date2 " & _
    " PIVOT auxiliary_statuses.REASON2;"
    qd.sql = sql
    qd.Close: Set qd = Nothing
    db.Close: Set db = Nothing

    'REFRESH PRODUCTION REPORT, CLOSE ALL REPORTS IF OPEN. MUST BE CLOSED BEFORE EXPORT
    DoCmd.OpenReport "Production", acViewNormal, "", "", acHidden
    DoCmd.Close acReport, "Production", acSaveYes
    If Application.CurrentProject.AllReports("Aux Status").IsLoaded = True Then
        DoCmd.Close acReport, "Aux Status", acSaveYes
    End If
    If Application.CurrentProject.AllReports("Production").IsLoaded = True Then
        DoCmd.Close acReport, "Production", acSaveYes
    End If
    If Application.CurrentProject.AllReports("Review Items").IsLoaded = True Then
        DoCmd.Close acReport, "Review Items", acSaveYes
    End If
    If Application.CurrentProject.AllReports("Scorecard").IsLoaded = True Then
        DoCmd.Close acReport, "Scorecard", acSaveYes
    End If
    
    'update coaching_log with work_status_id
    DoCmd.runSQL ("UPDATE coaching_log SET scorecard_printed = 1 WHERE (week_id = " & cboWeekOf.Value & " AND employee_id = " & cboSpecialist.Value & ");")

    'DBEngine.Idle dbRefreshCache
    Call ExportScorecard(specialistname)
    End
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Home; btnPrint_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select

End Sub
Private Sub ExportScorecard(specialistname As String)
    
    'EXPORT TO PDF
    Dim timeStamp As String: timeStamp = Format(Now(), "yyyy_mm_dd_ss")
    Dim filePath As String
    
    filePath = "X:\Member Enrollment\Member Enrollment(Custom)\Marketplace\Team Quality- DO NOT TOUCH\QA Database\Score cards\" & specialistname & "_" & timeStamp & ".pdf"
    DoCmd.SetWarnings False
    DoCmd.OutputTo acOutputReport, "Scorecard", acFormatPDF, filePath
    
    If MsgBox("Printing complete. Would you like to view the exported scorecard?", vbYesNo) = vbYes Then
        Shell "explorer.exe " & "X:\Member Enrollment\Member Enrollment(Custom)\Marketplace\Team Quality- DO NOT TOUCH\QA Database\Score cards\", vbNormalFocus
    End If

End Sub
Private Sub btnCoachingReports_Click()
    Select Case optCoachingReports
        Case 1
            DoCmd.OpenQuery ("qryCoachingLog")
        Case 2
            DoCmd.OpenQuery ("qryMonthlyItemsReviewed")
        Case 3
            DoCmd.OpenQuery ("qryTrainingLog")
    End Select
End Sub
Private Sub btnWeeklyReports_Click()
    Select Case optWeeklyReports
        Case 1
            DoCmd.OpenQuery "qryWeeklyErrors", acViewNormal, acReadOnly
        Case 2
            DoCmd.OpenQuery "qryWeeklyTeamScores", acViewNormal, acReadOnly
        Case 3
            DoCmd.OpenQuery "qryWeeklyTeamAvg", acViewNormal, acReadOnly
        Case 4
            DoCmd.OpenQuery "qryWeeklyBelow98", acViewNormal, acReadOnly
        Case 5
            DoCmd.OpenQuery "qryWeeklyErrorCount", acViewNormal, acReadOnly
        Case 6
            DoCmd.OpenQuery "qryWeeklyErrorCountByProcess", acViewNormal, acReadOnly
        Case 7
            Dim category(1 To 8) As String
            category(1) = "qryHICS"
            category(2) = "qryOnTermJob"
            category(3) = "qryOffTermJob"
            category(4) = "qryCancelJob"
            category(5) = "qryOther"
            category(6) = "qryRecon"
            category(7) = "qryCutlogDemoChanges"
            category(8) = "qryChat"
            
            For i = 1 To 8 Step 1
                DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, category(i), "X:\Member Enrollment\Member Enrollment(Custom)\Marketplace\Team Quality- DO NOT TOUCH\QA Database\Weekly Reporting\Team Scores_" & cboWeekOf2.Column(1) & ".xlsx", 1
            Next i
            
            If MsgBox("Complete. Would you like to view the report?", vbYesNo) = vbYes Then
             Shell "explorer.exe " & "X:\Member Enrollment\Member Enrollment(Custom)\Marketplace\Team Quality- DO NOT TOUCH\QA Database\Weekly Reporting\", vbNormalFocus
            Else: Exit Sub
            End If
    End Select
End Sub
Private Sub btnViewAdHoc_Click()
    Select Case optAdHoc
        Case 1
            DoCmd.OpenQuery "qryAdHocScores", acViewNormal, acReadOnly
        Case 2
            DoCmd.OpenQuery "qryAdHocProduction", acViewNormal, acReadOnly
        Case 3
            DoCmd.OpenQuery "qryAdHocErrors", acViewNormal, acReadOnly
        Case 4
            DoCmd.OpenQuery "qryAdHocAllItems", acViewNormal, acReadOnly
    End Select
End Sub

Private Sub optAdHoc_Click()
    If IsNull(cboStart.Value) Then
        cboStart.Value = getCurrentWeekId()
        cboEnd.Value = getCurrentWeekId()
    End If
End Sub

Private Sub Option30_GotFocus()
    cboMonth.Value = Format(Now(), "mmmm")
    cboYear.Value = Format(Now(), "yyyy")
End Sub
Private Sub optWeeklyReports_Click()
    If cboWeekOf2.Value = "" Or IsNull(cboWeekOf2) Then
        cboWeekOf2.Value = (getCurrentWeekId() - 1)
    End If
End Sub
Private Sub txtAreaOfExcellence_Exit(Cancel As Integer)
    If txtAreaOfExcellence.Value = "" Or IsNull(txtAreaOfExcellence) Then
        txtAreaOfExcellence.ForeColor = RGB(166, 166, 166)
    Else
        txtAreaOfExcellence.ForeColor = vbBlack
    End If
End Sub
Private Sub txtAreaOfExcellence_KeyDown(KeyCode As Integer, Shift As Integer)
    txtAreaOfExcellence.ForeColor = vbBlack
End Sub

Private Sub Label46_Click()
    DoCmd.OpenForm "QA Input", acNormal, , , , acWindowNormal, userId
End Sub

Private Sub Label44_Click()
    DoCmd.OpenForm "Record Locator", acNormal, , , , acWindowNormal, ",," & userId
End Sub

Private Sub Label51_Click()
    DoCmd.OpenForm "Admin", acNormal, , , acFormPropertySettings
End Sub

Private Sub Label58_Click()
    DoCmd.OpenForm "Chats", acNormal, , , acFormPropertySettings
End Sub
Private Sub Label67_Click()
    DoCmd.OpenForm "PROMPT", acNormal, , , , acFormPropertySettings, userId
End Sub
Private Sub Label79_Click()
    DoCmd.OpenForm "SysInfo", acNormal, , , , acFormPropertySettings, userId
End Sub
Public Function ValidateForm(submitType As Integer) As Boolean

    Dim msgStr As String, headerStr As String, footerStr As String, ctlName As String
    Dim varFocus As Variant
    
    headerStr = "<ul>"
    footerStr = "</ul>"
    
    Select Case submitType
        Case 1
            If cboWeekOf.Value = "" Or IsNull(cboWeekOf.Value) = True Then
                msgStr = msgStr & "<li><b>Week Of</b> cannot be blank.</li>"
                ctlName = "cboWeekOf,"
            End If
            If cboSpecialist.Value = "" Or IsNull(cboSpecialist.Value) = True Then
                msgStr = msgStr & "<li><b>Specialist</b> cannot be blank.</li>"
                ctlName = ctlName & "cboSpecialist,"
            End If
            If txtAreaOfExcellence.Value = "" Or IsNull(txtAreaOfExcellence.Value) = True Then
                msgStr = msgStr & "<li><b>Area of Excellence</b> cannot be blank.</li>"
                ctlName = ctlName & "txtAreaOfExcellence,"
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
Public Sub ClearForm()
    txtAreaOfExcellence.Value = Null
    txtAreaOfExcellence.ForeColor = RGB(166, 166, 166)
    txtEmployeeComments.Value = Null
    txtErrorBar.Value = Null
    txtErrorBar.BackColor = RGB(245, 246, 246)
    txtErrorBox.Value = Null
End Sub
