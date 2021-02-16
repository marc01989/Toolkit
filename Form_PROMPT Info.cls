VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PROMPT Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim varPriorNotes As Variant
Dim recordMax As Long, recordCounter As Long


Private Sub btnSearch_Click()
    
    If ValidateForm(1) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If
    Call ClearForm
    
On Error GoTo err1:

    '----SEARCH INFO HERE----
    Dim productionStr As String, qualityStr As String, noteStr As String, priorNotesStr As String, test As Variant
    Dim weekId As Integer, weekStart As Integer, employeeId As Integer
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    employeeId = cboSpecialist.Value
    weekId = getCurrentWeekId
    weekStart = weekId - 26

    
    'run query to get avg production data and concat results into string
    Set rs = db.OpenRecordset("TRANSFORM ROUND(AVG(weekly_production.amount), 0) AS AvgOfamount " & _
     " SELECT employees.employee_name " & _
     " FROM (production_segments INNER JOIN (weekly_production INNER JOIN employees ON weekly_production.employee_id = employees.employee_id) ON production_segments.segment_id = weekly_production.segment_id) INNER JOIN weeks ON weekly_production.week_id = weeks.week_id " & _
     " WHERE (weekly_production.employee_id = " & employeeId & " AND weekly_production.week_id > " & weekStart & ") " & _
     " GROUP BY  employees.employee_name, weekly_production.employee_id " & _
     " PIVOT production_segments.segment;", dbOpenSnapshot)
        'if returns records, loop and put into a string w/ bold and break tags. txt6MonthProductionAvg is rich text
        With rs
            If .recordCount > 0 Then
                For i = 1 To 10     'start i at 1 to skip the name field (data is redundant)
                productionStr = productionStr & "<b>" & .Fields(i).Name & "</b>: " & vbTab & .Fields(i) & "<br>"
                Next i
            End If
            txt6MonthProductionAvg.Value = productionStr
        End With
        rs.Close
    
    'query for member's 6 month score averages
    Set rs = db.OpenRecordset("SELECT employees.employee_name, Round(Avg(hics),2) AS [HICS Avg], Round(Avg(on_term_job),2) AS [On Term Job], Round(Avg(off_term_job),2) AS [Off Term Job], Round(Avg(cancel_job),2) AS [Cancel Job], Round(Avg(other),2) AS [Reporting], Round(Avg(recon),2) AS [Recon Avg], Round(Avg(cutlog_demo_changes),2) AS [Cut Demo Change], Round(Avg(chat),2) AS [Chats], Round(Avg(external_qa),2) AS [External QA] " & _
     " FROM employees INNER JOIN scores_mkt ON employees.employee_id = scores_mkt.employee_id " & _
     " WHERE (scores_mkt.week_id > " & weekStart & " And scores_mkt.employee_id = " & employeeId & ") " & _
     " GROUP BY employees.employee_name", dbOpenSnapshot)
        'if returns records, loop and put into a string w/ bold and break tags. txt6MonthProductionAvg is rich text
        With rs
            If .recordCount > 0 Then
                For i = 1 To 9      'start i at 1 to skip the name field (data is redundant)
                qualityStr = qualityStr & "<b>" & .Fields(i).Name & "</b>: " & vbTab & .Fields(i) & "<br>"
                Next i
            End If
            txt6MonthQualityAvg.Value = qualityStr
        End With
        rs.Close
    
    'query for basic employee info
    Set rs = db.OpenRecordset("SELECT employees.hire_date, employees.last_promotion_date, employees.is_new_hire, employees.title, employees.employee_number, employees_1.employee_name " & _
    " FROM employees INNER JOIN employees AS employees_1 ON employees.coach_id = employees_1.employee_id " & _
    " WHERE employees.[employee_id]= " & employeeId & ";", dbOpenSnapshot)
        With rs
            If .recordCount > 0 Then
                txtTitle.Value = rs.Fields(3)
                txtHireDate.Value = rs.Fields(0)
                Call CalculatePromotion(rs.Fields(1), rs.Fields(0), rs.Fields(2))
                txtEmployeeNumber.Value = rs.Fields(4)
                txtCoach.Value = rs.Fields(5)
            End If
        End With
        rs.Close

    'get that last review date and send to form
    Set rs = db.OpenRecordset("SELECT TOP 1, review_date " & _
    " FROM employee_review WHERE employee_id = " & employeeId & " ORDER BY ID DESC;")
        With rs
            If .recordCount > 0 Then
                test = Split(rs![review_date], " ")
                txtLastReview.Value = test(0)
            End If
        End With
        rs.Close
        
    'specialist career pathing comments
    Set rs = db.OpenRecordset("SELECT submit_date AS [Submit Date], comment AS [Specialist Notes] " & _
    " FROM career_pathing_comments WHERE employee_id = " & employeeId & " AND status = 1 " & _
    " ORDER BY ID;", dbOpenSnapshot)
    With rs
        If .recordCount > 0 Then
            .MoveFirst
            .MoveLast
            .MoveFirst
            Do While Not .EOF
                notesStr = notesStr & "<b>Submit Date: </b>" & .Fields(0).Value & "<br> <b>Notes: </b>" & .Fields(1).Value & "<br><br>"
                .MoveNext
            Loop
        End If
    End With
    rs.Close
    txtSpecialistCareerPathingNotes = notesStr
    noteStr = ""
    
    'prior meeting notes
    Set rs = db.OpenRecordset("SELECT ID, review_date, meeting_notes, projects_of_interest, career_goals, ideas_improvements, trends, concerns, achievements " & _
    " FROM employee_review " & _
    " WHERE employee_id = " & employeeId & " ORDER BY ID DESC;")
    With rs
        If .recordCount > 0 Then
            .MoveFirst
            .MoveLast
            .MoveFirst
            recordMax = .recordCount - 1
            recordCounter = 0
            Do While Not .EOF
                priorNotesStr = priorNotesStr & "<b>Date: </b>" & .Fields(1).Value & "<br><br><b>Meeting Notes: </b>" & .Fields(2).Value & _
                "<br><br><b>Projects of Interest: </b>" & .Fields(3).Value & "<br><br><b>Career Goals: </b>" & .Fields(4).Value & _
                "<br><br><b>Ideas for Improvements: </b>" & .Fields(5).Value & "<br><br><b>Trends: </b>" & .Fields(6).Value & _
                "<br><br><b>Achievements: </b>" & .Fields(7).Value & "|"
                .MoveNext
            Loop
            varPriorNotes = Split(Left(priorNotesStr, Len(priorNotesStr) - 1), "|")
            txtPriorMeetingNotes.Value = varPriorNotes(recordCounter)
        Else
            arrayChats = Empty
            recordMax = 0
            recordCounter = 0
        End If
    End With
    
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    
    GetWorkDist (employeeId)
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "PROMPT Info; btnSearch_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select

    
End Sub

Public Sub ClearForm()
    
    txt6MonthProductionAvg.Value = Null
    txt6MonthQualityAvg.Value = Null
    txtTitle.Value = Null
    txtHireDate.Value = Null
    txtLastPromotion.Value = Null
    txtEmployeeNumber.Value = Null
    txtCoach.Value = Null
    txtEligibleForPromotion.Value = Null
    txtMeetingNotes = Null
    txtCareerGoals = Null
    txtProjectsOfInterest = Null
    txtIdeasForImprovements = Null
    txtTrends = Null
    txtConcerns = Null
    txtSuccessesAndAchievements = Null
    txtNextPromotion = Null
    txtLastReview = Null
    txtPriorMeetingNotes = Null
    txtSpecialistCareerPathingNotes = Null
    lstWorkDistribution.RowSource = ""
    varPriorNotes = Empty
    'lstTrainingOpportunities.RowSource = ""
    'cboTrainingOpportunities.Value = ""
    'txtTrainingOther.Value = Null

End Sub
Private Sub btnForward_Click()
    If recordMax = 0 Then Exit Sub
    If recordCounter = recordMax Then
        Exit Sub
    Else
        recordCounter = recordCounter + 1
        txtPriorMeetingNotes.Value = varPriorNotes(recordCounter)
    End If
End Sub
Private Sub btnBack_Click()
    If recordMax = 0 Then Exit Sub
    If recordCounter = 0 Then
        Exit Sub
    Else
        recordCounter = recordCounter - 1
        txtPriorMeetingNotes.Value = varPriorNotes(recordCounter)
    End If
End Sub
Private Sub cboSpecialist_Change()
    Call ClearForm
End Sub
Private Sub Label12_Click()
    DoCmd.Close acForm, "PROMPT Info", acSaveNo
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
        Case 1
            If cboSpecialist.Value = "" Or IsNull(cboSpecialist.Value) = True Then
                msgStr = msgStr & "<li><b>Specialist</b> cannot be blank.</li>"
                ctlName = ctlName & "cboSpecialist,"
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

Public Sub CalculatePromotion(lastPromotionDate As Variant, hireDate As Variant, isNewHire As Variant)
    Dim nextPromoDate As String, currentDate As String
    
    If isNewHire = 1 Then 'new hire, specialist hasn't worked 1 year in the dept yet
        nextPromoDate = Format(DateAdd("d", 365, hireDate), "yyyymmdd")
        currentDate = Format(Now(), "yyyymmdd")
        txtLastPromotion.Value = "New Hire"
        txtNextPromotion.Value = DateAdd("d", 365, hireDate)
        
        If currentDate >= nextPromoDate Then
            txtEligibleForPromotion = "True"
        Else
            txtEligibleForPromotion = "False"
        End If
        
    Else 'specialist has been promoted before, aka not a new hire
        nextPromoDate = Format(DateAdd("d", 182, lastPromotionDate), "yyyymmdd")
        currentDate = Format(Now(), "yyyymmdd")
        txtLastPromotion.Value = lastPromotionDate
        txtNextPromotion = DateAdd("d", 182, lastPromotionDate)
        
        If currentDate >= nextPromoDate Then
            txtEligibleForPromotion = "True"
        Else
            txtEligibleForPromotion = "False"
        End If
    End If

End Sub

Public Sub GetWorkDist(employeeId As Integer)

    Dim conn As ADODB.Connection
    Dim rsa As ADODB.Recordset
    Dim strCon As String, sql As String
    
    strCon = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source=X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\iWD\iWD database.accdb;" & _
    "User Id=admin;Password="

    'get work distribution, send to listbox
    sql = "SELECT skill.skill_name AS Skill, dist_status.dist_status AS Distribution, scorecard_category.scorecard_category AS Category, process.frequency " & _
    " FROM (specialist INNER JOIN (((process INNER JOIN scorecard_category ON process.scorecard_category_id = scorecard_category.scorecard_category_id) INNER JOIN skill ON process.process_id = skill.process_id) INNER JOIN dist_junction ON skill.skill_id = dist_junction.skill_id) ON specialist.employee_id = dist_junction.specialist_id) INNER JOIN dist_status ON dist_junction.dist_status_id = dist_status.dist_status_id " & _
    " WHERE (((Specialist.[employee_id]) = " & employeeId & ")) " & _
    " ORDER BY scorecard_category, process.process_id, skill.skill_id;"
    
On Error GoTo err1:

    '--START CONNECTION HERE--
    Set conn = New ADODB.Connection
    Set rsa = New ADODB.Recordset
    conn.Open (strCon)
    rsa.Open sql, conn, adOpenKeyset
    
    With rsa
        If .recordCount > 0 Then
            .MoveLast
            .MoveFirst
                lstWorkDistribution.AddItem .Fields(0).Name & ";" & .Fields(1).Name & ";" & .Fields(2).Name & ";" & .Fields(3).Name
            Do While Not .EOF
                lstWorkDistribution.AddItem .Fields(0).Value & ";" & .Fields(1).Value & ";" & .Fields(2).Value & ";" & .Fields(3).Value
            .MoveNext
            Loop
        End If
    End With
    
    rsa.Close: Set rsa = Nothing
    conn.Close: Set conn = Nothing
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "PROMPT Info; GetWorkDist()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
End Sub
