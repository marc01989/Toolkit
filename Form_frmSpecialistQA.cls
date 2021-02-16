VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSpecialistQA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim weekId As Integer, employeeId As Integer, currentWeekId As Integer
Dim arrayWeekScore As Variant, arrayErrors As Variant, arrayProduction As Variant, arrayComments As Variant
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim strCon As String, weekName As String

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'this form show specialist their QA scores by using ADO OLEDB connection to QA database
'using this type of connection is useful bc it allows user to access database w/o perminantly linking to all the db's tables
'data is queries upon loading of form
'all data is queried and stored into local arrays
'-----------------------------------------------------------------------------------------------------------------------------------------------------------

'--REFERENCES NEEDED FOR ADO CONNECTION
    '-Microsoft Activex Data Objects 6.0 Library
    

Private Sub Form_Close()
    '--KILL CONNECTION------
    Set rs = Nothing
    conn.Close: Set conn = Nothing
    '-----------------------
End Sub
Private Sub Form_Load()

On Error GoTo err1:
    '--START CONNECTION HERE--
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    strCon = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source=X:\Member Enrollment\Member Enrollment(Custom)\Marketplace\Database\QA Database Backend\QA_Tracker_BE.accdb;" & _
    "User Id=admin;Password="
    conn.Open (strCon)
    '-------------------------
    
    Dim today As String, networkID As String, sql As String
    Dim wkd As Integer
    wkd = Weekday(Date)
    wkd = wkd - 1
    today = Date - wkd
    today = Format(today, "mmmm d, yyyy")
    networkID = Environ$("username")
    
    '--OPEN CONNECTIONS HERE--
    'get employee_id by querying employees table
    sql = "SELECT employee_id FROM employees WHERE network_id = '" & networkID & "';"
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    With rs
        If .recordCount > 0 Then
            .MoveFirst
            .MoveLast
            .MoveFirst
            employeeId = .Fields(0)
        Else
            If MsgBox("Specialist not found", vbExclamation + vbOKOnly) = vbOK Then
                DoCmd.Close acForm, "frmSpecialistQA", acSaveNo
                 Exit Sub
            End If
        End If
    End With
    rs.Close


    'get current week_id by querying weeks table
    sql = "SELECT week_id, week_start FROM weeks WHERE week_start = '" & today & " ';"
    rs.Open sql, conn, adOpenKeyset
    With rs
        If .recordCount > 0 Then
            .MoveFirst
            .MoveLast
            .MoveFirst
            weekId = .Fields(0)
            currentWeekId = weekId
            weekName = .Fields(1)
            txtWeekOf = weekName
        End If
    End With
    rs.Close
    
    'queries all weekly scores for emp and saves in array
    sql = "SELECT weeks.week_start, scores_mkt.week_id, scores_mkt.HICs, scores_mkt.ON_term_job, scores_mkt.OFF_term_job, scores_mkt.cancel_job, scores_mkt.other, scores_mkt.recon, scores_mkt.cutlog_demo_changes, scores_mkt.chat " & _
         " FROM employees INNER JOIN (weeks INNER JOIN scores_mkt ON weeks.week_id = scores_mkt.week_id) ON employees.employee_id = scores_mkt.employee_id " & _
         " WHERE (scores_mkt.employee_id = " & employeeId & " );"
    rs.Open sql, conn, adOpenKeyset
    
    With rs
        If .recordCount > 0 Then
            .MoveFirst
            .MoveLast
            .MoveFirst
            arrayWeekScore = .GetRows(.recordCount)
        End If
    End With
    rs.Close
    
    'queries all errors for emp and saves in array
    sql = "SELECT week_id, member_id AS [Member ID], additional_id AS [Other ID], error_name AS [Error Category], error_txt AS [Description]" & _
         " FROM errors INNER JOIN review_items_mkt ON errors.error_id = review_items_mkt.error_id " & _
         " WHERE (employee_id = " & employeeId & " AND review_items_mkt.error_id <> 0);"
    rs.Open sql, conn, adOpenKeyset
    
    With rs
        If .recordCount > 0 Then
            .MoveFirst
            .MoveLast
            .MoveFirst
            arrayErrors = .GetRows(.recordCount)
        End If
    End With
    rs.Close
    
    'queries all production for emp and saves in array
    sql = "TRANSFORM Sum(weekly_production.amount) AS SumOfamount " & _
         " SELECT weeks.week_id, employees.employee_name " & _
         " FROM (production_segments INNER JOIN (weekly_production INNER JOIN employees ON weekly_production.employee_id = employees.employee_id) ON production_segments.segment_id = weekly_production.segment_id) INNER JOIN weeks ON weekly_production.week_id = weeks.week_id " & _
         " WHERE (weekly_production.employee_id = " & employeeId & ") " & _
         " GROUP BY weeks.week_id, employees.employee_name, weekly_production.employee_id " & _
         " ORDER BY weeks.week_id " & _
         " PIVOT production_segments.segment;"
    rs.Open sql, conn, adOpenKeyset
    
    With rs
        If .recordCount > 0 Then
            .MoveFirst
            .MoveLast
            .MoveFirst
            arrayProduction = .GetRows(.recordCount)
        End If
    End With
    rs.Close
    
    'queries all records where specialist is on coaching log, independent of whether or not coaching/careerpathing comments are null
    sql = "SELECT coaching_log.week_id, employee_comments, comment " & _
         " FROM career_pathing_comments RIGHT JOIN coaching_log ON (career_pathing_comments.employee_id = coaching_log.employee_id) AND (career_pathing_comments.week_id = coaching_log.week_id) " & _
         " WHERE coaching_log.employee_id = " & employeeId & ";"
    rs.Open sql, conn, adOpenKeyset
    
    With rs
        If .recordCount > 0 Then
            .MoveFirst
            .MoveLast
            .MoveFirst
            arrayComments = .GetRows(.recordCount)
        End If
    End With
    rs.Close: Set rs = Nothing
    
    'call subs to moving through arrays depending on weekID
    Call showAvgScore(weekId, employeeId)
    Call showWeekScore(weekId)
    Call showErrors(weekId)
    Call showProduction(weekId)
    Call showComments(weekId)

err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmSpecialistQA; Form_Load()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select

End Sub

Public Function showAvgScore(weekId As Integer, employeeId As Integer)

    'calculates the avg scores when user iterates through weeks
    diff = weekId - 26
    Set rs = New ADODB.Recordset
    sql = "SELECT ROUND(Avg(HICs), 2) AS HICsAvg, ROUND(Avg(ON_term_job), 2) AS OnTermAvg, ROUND(Avg(OFF_term_job), 2) AS OffTermAvg, " & _
        " ROUND(AVG(cancel_job), 2) AS CancelJobAvg, ROUND(AVG(other), 2) AS ReportingAvg, ROUND(AVG(recon), 2) AS ReconAvg, ROUND(AVG(cutlog_demo_changes), 2) AS CutDemoAvg, ROUND(AVG(chat), 2) AS ChatAvg, ROUND(AVG(external_qa), 2) " & _
        " FROM scores_mkt " & _
        " WHERE [employee_id]= " & employeeId & " AND [week_id]<= " & weekId & " And [week_id]>= " & diff & ";"
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    With rs
        txtHICsAvg.Value = rs![HICsAvg]
        txtOnTermAvg.Value = rs![OnTermAvg]
        txtOffTermAvg.Value = rs![OffTermAvg]
        txtCancelJobAvg.Value = rs![CancelJobAvg]
        txtReportingAvg.Value = rs![ReportingAvg]
        txtReconAvg.Value = rs![ReconAvg]
        txtCutDemoAvg.Value = rs![CutDemoAvg]
        txtChatAvg.Value = rs![ChatAvg]
    End With
    rs.Close: Set rs = Nothing
    
End Function

Public Function showWeekScore(weekId As Integer)
'shows week scores

If Not IsEmpty(arrayWeekScore) Then
    For i = 0 To UBound(arrayWeekScore, 2)
        If arrayWeekScore(1, i) = weekId Then
            txtHICs.Value = arrayWeekScore(2, i)
            txtOnTerm.Value = arrayWeekScore(3, i)
            txtOffTerm.Value = arrayWeekScore(4, i)
            txtCancelJob.Value = arrayWeekScore(5, i)
            txtReporting.Value = arrayWeekScore(6, i)
            txtRecon.Value = arrayWeekScore(7, i)
            txtCutDemo.Value = arrayWeekScore(8, i)
            txtChat.Value = arrayWeekScore(9, i)
            Exit For
        Else
            txtHICs.Value = Null
            txtOnTerm.Value = Null
            txtOffTerm.Value = Null
            txtCancelJob.Value = Null
            txtReporting.Value = Null
            txtRecon.Value = Null
            txtCutDemo.Value = Null
            txtChat.Value = Null
        End If
    Next i
End If
End Function

Public Function showErrors(weekId As Integer)
'shows errors for week

    Dim errorText As String: errorText = ""
    
    If Not IsEmpty(arrayErrors) Then
        For i = 0 To UBound(arrayErrors, 2)
            If arrayErrors(0, i) = weekId Then
                errorText = errorText & "<b>Member ID: </b>" & arrayErrors(1, i) & "<br>" & _
                "<b>Other ID</b>: " & arrayErrors(2, i) & "<br>" & _
                "<b>Error Category: </b>" & arrayErrors(3, i) & "<br>" & _
                "<b>Description: </b>" & arrayErrors(4, i) & "<br><br>"
            End If
        Next i
    End If
    txtErrors.Value = errorText

End Function

Public Function showProduction(weekId As Integer)
'shows production for week

If Not IsEmpty(arrayProduction) Then
    For i = 0 To UBound(arrayProduction, 2)
        If arrayProduction(0, i) = weekId Then
            txtChats.Value = arrayProduction(2, i)
            txtCutlogs.Value = arrayProduction(3, i)
            txtEmails.Value = arrayProduction(4, i)
            txtHICsClosed.Value = arrayProduction(5, i)
            txtHICsOutreach.Value = arrayProduction(6, i)
            txtInboundCalls.Value = arrayProduction(7, i)
            txtOutboundCalls.Value = arrayProduction(8, i)
            txtReports.Value = arrayProduction(9, i)
            txtScans.Value = arrayProduction(10, i)
            txtSelfReporting.Value = arrayProduction(11, i)
            Exit For
        Else
            txtCutlogs.Value = Null
            txtEmails.Value = Null
            txtChats.Value = Null
            txtHICsClosed.Value = Null
            txtHICsOutreach.Value = Null
            txtInboundCalls.Value = Null
            txtOutboundCalls.Value = Null
            txtReports.Value = Null
            txtScans.Value = Null
            txtSelfReporting.Value = Null
        End If
    Next i
End If

End Function

Public Function showComments(weekId As Integer)
'shows any comments user submitted for a week

If Not IsEmpty(arrayComments) Then
    For i = 0 To UBound(arrayComments, 2)
        If arrayComments(0, i) = weekId Then
            txtCoachingConcerns.Value = arrayComments(1, i)
            txtCareerPathConcerns.Value = arrayComments(2, i)
            Exit For
        Else
            txtCoachingConcerns.Value = Null
            txtCareerPathConcerns.Value = Null
        End If
    Next i
End If

End Function

Private Sub btnBack_Click()
'previous week's records

    weekId = weekId - 1
    If weekId > currentWeekId Then
        MsgBox ("Cannot go past current week")
    Else
        weekName = DateAdd("d", -7, weekName)
        txtWeekOf.Value = weekName
        Call showAvgScore(weekId, employeeId)
        Call showWeekScore(weekId)
        Call showErrors(weekId)
        Call showProduction(weekId)
        Call showComments(weekId)
    End If
End Sub

Private Sub btnForward_Click()
'next weeks records

    weekId = weekId + 1
    If weekId > currentWeekId Then
        weekId = weekId - 1
        MsgBox ("Cannot go past current week")
    Else
        weekName = DateAdd("d", 7, weekName)
        txtWeekOf.Value = weekName
        Call showAvgScore(weekId, employeeId)
        Call showWeekScore(weekId)
        Call showErrors(weekId)
        Call showProduction(weekId)
        Call showComments(weekId)
    End If
End Sub

Private Sub btnSubmitCoachingConcerns_Click()
'submit comments for week

    '--VALIDATION--
    If IsNull(cboProcessType) Or cboProcessType.Value = "" Then
        If MsgBox("You must select a 'Process Type' before submitting.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then: Exit Sub
    End If

On Error GoTo err1:
    
    '--START CONNECTION/SUBMISSION--
    If MsgBox("Submit a coaching comment for the week of " & txtWeekOf.Value & "?", vbQuestion + vbYesNo) = vbYes Then
        Set rs = New ADODB.Recordset
        Dim sql As String
        sql = "UPDATE coaching_log SET employee_comments = '" & cleanText(txtCoachingConcerns.Value) & "'," & _
        " process_id = '" & cboProcessType.Value & "'" & _
        " WHERE week_id = " & weekId & " AND employee_id = " & employeeId & ";"
        rs.Open "SELECT * FROM coaching_log", conn, adOpenKeyset
        conn.Execute sql
    'update array value
        For i = 0 To UBound(arrayComments, 2)
            If arrayComments(0, i) = weekId Then arrayComments(1, i) = txtCoachingConcerns.Value: Exit For
        Next i
        
        cboProcessType.Value = Null
        rs.Close: Set rs = Nothing
    End If

    MsgBox ("Comment submitted")
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmSpecialistQA; btnSubmitCoachingConcerns_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
End Sub

Private Sub btnSubmitCareerPathing_Click()
    '--VALIDATION--
    If IsNull(txtCareerPathConcerns) Or txtCareerPathConcerns.Value = "" Then
        If MsgBox("'Career Pathing' notebox cannot be blank.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then: Exit Sub
    End If
    
On Error GoTo err1:
    
    '--START CONNECTION/SUBMISSION--
    If MsgBox("Submit career pathing notes for the week of " & txtWeekOf.Value & "?", vbQuestion + vbYesNo) = vbYes Then
        Set rs = New ADODB.Recordset
        Dim sql As String
        sql = "INSERT INTO career_pathing_comments (week_id, employee_id, comment, submit_date, status) " & _
        " VALUES (" & weekId & ", " & employeeId & ", '" & cleanText(txtCareerPathConcerns.Value) & "', '" & Now() & "', 1);"
        'rs.Open "SELECT * FROM coaching_log", conn, adOpenKeyset
        Debug.Print sql
        conn.Execute sql
    'update array value
        For i = 0 To UBound(arrayComments, 2)
            If arrayComments(0, i) = weekId Then arrayComments(2, i) = txtCareerPathConcerns.Value: Exit For
        Next i
    End If

err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmSpecialistQA; btnSubmitCareerPathing_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select

End Sub


