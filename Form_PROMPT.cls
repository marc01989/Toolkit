VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_PROMPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim userId As Integer

Private Sub Form_Load()

    Dim obj As AccessObject, dbs As Object
    Dim args As String
    Set db = CurrentDb
    
    'checks if value is passed when opening form, runs main query if so
    If Not IsNull(Forms![PROMPT].OpenArgs) Then
        args = Forms![PROMPT].OpenArgs
            If IsNull(args) Or (args <> 7 And args <> 34 And args <> 35) Then
                MsgBox ("User not found - please login again")
                Set dbs = Application.CurrentProject
                For Each obj In dbs.AllForms
                    If obj.IsLoaded = True Then
                      DoCmd.Close acForm, obj.Name, acSaveNo
                    End If
                Next obj
                DoCmd.OpenForm "Login", acNormal, , , , acWindowNormal
                Call LogError(0 & " " & "User Id not found or not passed to home screen", "PROMPT; Form_Load()")
                Exit Sub
            Else
                userId = CInt(args)
                'MsgBox (userId)
                Call btnClear_Click
                PromotionDefault
                lstTrainingOpportunities.RowSource = ""
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

Private Sub cboSpecialist_Change()
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("employees")
    With rs
        .FindFirst ("employee_id = " & cboSpecialist.Value)
        If .NoMatch Then
            Call btnClear_Click
        Else
            txtTitle.Value = rs![Title]
            Call CalculatePromotion(rs![last_promotion_date], rs![hire_date], rs![is_new_hire])
        End If
    End With
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    
End Sub
Private Sub btnSubmit_Click()
        
    If ValidateForm(1) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If

On Error GoTo err1:

    '----SUBMIT INFO HERE----
    Dim reviewID As Long
    Dim trainingStr As String
    Dim employeeId As Integer: employeeId = cboSpecialist.Value
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("employee_review")
    
    DoCmd.SetWarnings False
    If MsgBox("Submit new record for " & cboSpecialist.Column(1) & "?", vbYesNo, "Submit Record?") = vbYes Then
        With rs
            .AddNew
                rs![employee_id] = employeeId
                rs![review_date] = Now()
                rs![eligible_for_promotion] = txtEligibleForPromotion.Value
                rs![meeting_notes] = txtMeetingNotes.Value
                rs![projects_of_interest] = txtProjectsOfInterest.Value
                rs![career_goals] = txtCareerGoals.Value
                rs![ideas_improvements] = txtIdeasForImprovements.Value
                rs![trends] = txtTrends.Value
                rs![concerns] = txtConcerns.Value
                rs![achievements] = txtSuccessesAndAchievements.Value
                rs![submitter_id] = userId
                
                If cboPromotionAccepted.Value = "TRUE" Then
                    rs![promoted] = cboPromotionAccepted.Value
                    rs![Title] = txtNewTitle.Value
                    rs![promotion_date] = Format(dtpPromotionEffectiveDate.Value, "mm/dd/yyyy")
                    DoCmd.runSQL ("UPDATE employees " & _
                    " SET last_promotion_date = '" & Format(dtpPromotionEffectiveDate.Value, "mm/dd/yyyy") & "'" & _
                    " , title = '" & txtNewTitle & "' WHERE employee_id = " & employeeId & ";")
                Else
                    rs![Title] = txtTitle.Value
                End If
            .Update
            
            .Move 0, .LastModified
            reviewID = rs![Id]
            
            'update career_pathing table to show specialist notes are addressed/closed
            DoCmd.runSQL ("UPDATE career_pathing_comments SET status = 0, review_id = " & reviewID & " WHERE employee_id = " & employeeId & " AND status = 1;")
        End With
    Else
        Exit Sub
    End If
    rs.Close
    
    'test if training opportunities lstBox is null or not
    If lstTrainingOpportunities.ListCount > 0 Then
        Set rs = db.OpenRecordset("training_opportunities")
        For i = 0 To lstTrainingOpportunities.ListCount - 1
            With rs
                .AddNew
                rs![employee_id] = employeeId
                rs![submit_date] = Now()
                rs![Training] = lstTrainingOpportunities.ItemData(i)
                rs![review_id] = reviewID
                .Update
            End With
        Next i
    End If
    
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    
    Call btnClear_Click
    MsgBox ("Item submitted")

err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "PROMPT; btnSubmit_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select

End Sub
Private Sub cboPromotionOffered_Change()
    If cboPromotionOffered.Value = "TRUE" Then
        cboPromotionAccepted.Visible = True
    Else
        PromotionDefault
    End If
End Sub

Private Sub cboPromotionAccepted_Change()
    If cboPromotionAccepted.Value = "TRUE" Then
        txtNewTitle.Visible = True
        dtpPromotionEffectiveDate.Visible = True
    Else
        txtNewTitle.Visible = False
        dtpPromotionEffectiveDate.Visible = False
    End If
End Sub
Public Sub PromotionDefault()
    cboPromotionOffered.Value = "FALSE"
    cboPromotionAccepted.Value = "FALSE"
    cboPromotionAccepted.Visible = False
    txtNewTitle.Value = Null
    txtNewTitle.Visible = False
    'dtpPromotionEffectiveDate.Value = Now
    dtpPromotionEffectiveDate.Visible = False
End Sub
Private Sub btnAdd_Click()
    If txtTrainingOther.Value <> "" And Not IsNull(txtTrainingOther) Then
        lstTrainingOpportunities.AddItem ("Other: " & txtTrainingOther.Value)
        txtTrainingOther.Value = Null
    End If
End Sub

Private Sub btnRemove_Click()
    If lstTrainingOpportunities.ListIndex < 0 Then
    Else
        lstTrainingOpportunities.RemoveItem (lstTrainingOpportunities.ListIndex)
    End If
End Sub
Private Sub cboTrainingOpportunities_AfterUpdate()
    If cboTrainingOpportunities.Value <> "" And Not IsNull(cboTrainingOpportunities) Then
      lstTrainingOpportunities.AddItem (cboTrainingOpportunities.Column(1))
      cboTrainingOpportunities.Value = ""
    End If
End Sub
Private Sub btnClear_Click()
    txtMeetingNotes = Null
    txtCareerGoals = Null
    txtProjectsOfInterest = Null
    txtIdeasForImprovements = Null
    txtTrends = Null
    txtConcerns = Null
    txtSuccessesAndAchievements = Null
    txtNextPromotion = Null
    txtLastReview = Null
    txtSpecialistCareerPathingNotes = Null
    lstTrainingOpportunities.RowSource = ""
    cboTrainingOpportunities.Value = ""
    txtTrainingOther.Value = Null
    txtErrorBar.Value = Null
    txtErrorBar.BackColor = RGB(245, 246, 246)
    txtErrorBox.Value = Null
    txtTitle.Value = Null
    txtEligibleForPromotion.Value = Null
    cboSpecialist.Value = Null
    PromotionDefault
End Sub
Private Sub Label12_Click()
    DoCmd.Close acForm, "PROMPT", acSaveNo
End Sub

Private Sub Label46_Click()
    DoCmd.OpenForm "Home", acNormal, , , acFormPropertySettings
End Sub

Private Sub Label80_Click()
    DoCmd.OpenForm "PROMPT Info", acNormal, , , acFormPropertySettings
End Sub
Public Sub CalculatePromotion(lastPromotionDate As Variant, hireDate As Variant, isNewHire As Variant)
    Dim nextPromoDate As String, currentDate As String
    
    If isNewHire = 1 Then 'new hire, specialist hasn't worked 1 year in the dept yet
        nextPromoDate = Format(DateAdd("d", 365, hireDate), "yyyymmdd")
        currentDate = Format(Now(), "yyyymmdd")
        
        If currentDate >= nextPromoDate Then
            txtEligibleForPromotion = "True"
        Else
            txtEligibleForPromotion = "False"
        End If
        
    Else 'specialist has been promoted before, aka not a new hire
        nextPromoDate = Format(DateAdd("d", 182, lastPromotionDate), "yyyymmdd")
        currentDate = Format(Now(), "yyyymmdd")
        
        If currentDate >= nextPromoDate Then
            txtEligibleForPromotion = "True"
        Else
            txtEligibleForPromotion = "False"
        End If
    End If
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
            If cboPromotionAccepted.Value = "TRUE" And (IsNull(txtNewTitle) Or txtNewTitle.Value = "") Then
                msgStr = msgStr & "<li><b>New Title</b> cannot be blank when granting promotion.</li>"
                ctlName = ctlName & "txtNewTitle,"
            End If
            submitStr = txtMeetingNotes & txtCareerGoals & txtProjectsOfInterest & _
            txtIdeasForImprovements & txtTrends & txtConcerns & txtSuccessesAndAchievements
            If IsNull(submitStr) Or Len(submitStr) = 0 Then
                msgStr = msgStr & "<li>Submission cannot be blank. Please complete at least one note field.</li>"
                ctlName = ctlName & "txtMeetingNotes,"
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


