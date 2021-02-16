VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'This form houses a number of different reports and functions related mainly to HICs
'-----------------------------------------------------------------------------------------------------------------------------------------------------------


Private Sub btnMiscReports_Click()
    Select Case optMiscReports.Value
        Case 1
            DoCmd.OpenQuery "qry1095Disputes", acViewNormal, acReadOnly
    End Select
End Sub

Private Sub cboDate_AfterUpdate()
    cboDate2.Requery
End Sub

Private Sub cboDate_Click()
    cboDate2.Requery
End Sub
'################Sub tied to Generating AdHoc Uploads-- This searches and returns basic case data -- does not generate any reports############################
Private Sub cmdCaseSearch_Click()
Dim db As Database: Set db = CurrentDb
Dim rs As Recordset
Dim strSearch As String: strSearch = "[case_id] = " & "'" & txtCaseIDUpload.Value & "'"

    If Me.txtCaseIDUpload.Value <> "" And Not IsNull(Me.txtCaseIDUpload.Value) Then
    Set rs = db.OpenRecordset("tblCMSCaseData", dbOpenSnapshot)
    rs.FindFirst (strSearch)
        If rs.NoMatch Then
            rs.Close: Set rs = Nothing
            txtIssuerAssignDate.Value = ""
            txtComplianceDate.Value = ""
            txtCaseLevel.Value = ""
            MsgBox ("No Case Found")
        ElseIf Not rs.NoMatch Then
            txtIssuerAssignDate.Value = rs![issuer_assign_date]
            txtComplianceDate.Value = rs![compliance_due_date]
            txtCaseLevel.Value = rs![case_level]
            rs.Close: Set rs = Nothing
        End If
    Else: MsgBox ("Case ID cannot be empty")
    End If
    Set db = Nothing


End Sub
'Generates adhoc reports for upload into HICS system or into MC400 (cutlogs)
Private Sub cmdGenerateAdHocReport_Click()
    Dim db As Database: Set db = CurrentDb
    Dim rs As Recordset
    Dim filePath As String: exPath = "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\HICS DBs\UPLOADS\Ad_Hoc_Reports\"
    Dim strSQL As String:
    Dim strSearch As String: strSearch = "[case_id] = " & "'" & txtCaseIDUpload.Value & "'" & " And [uploaded] = 'N'"
'Checks to make sure there is actually a case in the CASEID field
    If txtCaseIDUpload.Value <> "" And Not IsNull(txtCaseIDUpload.Value) Then
            Select Case grpAdHocReport.Value
'IF the user selects Case Work Notes then this generates an AdHocCaseworkNotes Extract
            Case 1
                Set rs = db.OpenRecordset("tblCaseWorkNotes", dbOpenSnapshot)
                rs.FindFirst (strSearch)
                If rs.NoMatch Then
                    MsgBox ("No Casework Notes that have not been uploaded")
                ElseIf Not rs.NoMatch Then
                    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "qryHICSUploadUnresolvedAdHoc", exPath & "Ad_Hoc_Upload_Casework_Notes_" & txtCaseIDUpload.Value & "_" & Format(Now(), "mm_dd_yyyy_hhnnss")
                    DoCmd.OpenQuery "qupdCaseworkNotesAdHoc", acViewNormal, acEdit
                    If MsgBox("Go to Ad Hoc HICS Upload folder?", vbYesNo) = vbYes Then
                        Shell "explorer.exe " & exPath, vbNormalFocus
                    End If
                End If
                rs.Close: Set rs = Nothing
'If the usuer selects ResolutionNOtes then generates AdhocResolutionNotes Extract
            Case 2
                Set rs = db.OpenRecordset("tblCaseWorkNotes", dbOpenSnapshot)
                rs.FindFirst (strSearch)
                If rs.NoMatch Then
                    MsgBox ("No Resolution Notes that have not been uploaded")
                ElseIf Not rs.NoMatch Then
                    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "qryHICSUploadResolvedAdHoc", exPath & "Ad_Hoc_Upload_Resolution_Notes_" & txtCaseIDUpload.Value & "_" & Format(Now(), "mm_dd_yyyy_hhnnss")
                    DoCmd.OpenQuery "qupdResolutionNotesAdHoc", acViewNormal, acEdit
                    If MsgBox("Go to Ad Hoc HICS Upload folder?", vbYesNo) = vbYes Then
                        Shell "explorer.exe " & exPath, vbNormalFocus
                    End If
               End If
               rs.Close: Set rs = Nothing
' If the user selected CUTlog Notes then generates the AdHoc CutlogExtract
            Case 3
                Set rs = db.OpenRecordset("tblCutlogNotes", dbOpenSnapshot)
                rs.FindFirst (strSearch)
                If rs.NoMatch Then
                    MsgBox ("No Cutlog Notes that have not been uploaded")
                ElseIf Not rs.NoMatch Then
                    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "qryCutlogUploadAdHoc", exPath & "Ad_Hoc_Upload_Cutlog_Notes_" & txtCaseIDUpload.Value & "_" & Format(Now(), "mm_dd_yyyy_hhnnss")
                    strSQL = "UPDATE tblCutlogNotes SET tblCutlogNotes.extracted = 'Y', tblCutlogNotes.uploaded = 'Y', tblCutlogNotes.upload_date = Now() WHERE (((tblCutlogNotes.uploaded)='N') And tblCutlogNotes.case_id =" & "'" & txtCaseIDUpload.Value & "'" & ")"
                    db.Execute strSQL, dbFailOnError
                    If MsgBox("Go to Ad Hoc HICS Upload folder?", vbYesNo) = vbYes Then
                        Shell "explorer.exe " & exPath, vbNormalFocus
                    End If
                End If
                rs.Close: Set rs = Nothing
'If the user Selects ERR NOtes then this will generates ERRnotes as an AdHoc upload
                Case 4
                Set rs = db.OpenRecordset("tblERRNotes", dbOpenDynaset)
                strSearch = "[CaseID] = " & "'" & txtCaseIDUpload.Value & "'" & " And [Uploaded] = 0"
                rs.FindFirst (strSearch)
                If rs.NoMatch Then
                    MsgBox ("No ER&&R Notes that have not been uploaded")
                ElseIf Not rs.NoMatch Then
                    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "qryERRUploadAdHoc", exPath & "Ad_Hoc_Upload_ERR_Notes_" & txtCaseIDUpload.Value & "_" & Format(Now(), "mm_dd_yyyy_hhnnss")
                    strSQL = "UPDATE tblERRNotes SET tblERRNotes.extracted = 'Y', tblERRNotes.Uploaded = -1, tblERRNotes.UploadedDate = Now() WHERE (((tblERRNotes.Uploaded)= 0) And tblERRNotes.CaseID = " & "'" & txtCaseIDUpload.Value & "');"
                    db.Execute strSQL, dbFailOnError
                    If MsgBox("Go to AdHoc HICS Upload folder?", vbYesNo) = vbYes Then
                        Shell "explorer.exe " & exPath, vbNormalFocus
                    End If
                End If
                rs.Close: Set rs = Nothing
                
            End Select
    ElseIf txtCaseIDUpload.Value = "" Or IsNull(txtCaseIDUpload.Value) Then
                    MsgBox ("CaseID box cannot be Blank")
                    Exit Sub
    End If

        End Sub



Private Sub cmdGenerateReport_Click()
'Various reports in the HICS Reports section of the admin form
    Select Case grpReports.Value
    Case 1
        DoCmd.OpenQuery "qyrOpenHICSCasesCompare", acViewNormal, acReadOnly
    Case 2
        DoCmd.OpenQuery "qryResolvedCases", acViewNormal, acReadOnly
    Case 3
        DoCmd.OpenQuery "qryResolvedNotUploaded", acViewNormal, acReadOnly
    Case 4
        DoCmd.OpenQuery "qryERRCases", acViewNormal, acReadOnly
    Case 5
        DoCmd.OpenQuery "qryCasesWorkedToday", acViewNormal, acReadOnly
    Case 6
        DoCmd.OpenQuery "qryUnworkedCases", acViewNormal, acReadOnly
    Case 7
        DoCmd.OpenQuery "qryAdHocCaseNotes", acViewNormal, acReadOnly
    Case 8
        DoCmd.OpenQuery "qryFunctionalityLimitation", acViewNormal, acReadOnly
    Case 9
        DoCmd.OpenQuery "qryMisrepresentationCases", acViewNormal, acReadOnly
    Case 10
        DoCmd.OpenQuery "qryTransferredCases", acViewNormal, acReadOnly
    End Select
    
End Sub

Private Sub cmdGenerateReports_Click()

'Bound to "IWD Reports section" that generates all the reports needed for the morning load
    Dim crtDate As String
    crtDate = Format(Now(), "yyyymmdd")
    Select Case grpIWDReports.Value
        Case 1
            DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "qryMorningLoadNewCases", _
                    "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\HICS DBs\UPLOADS\" & "Morning_Reports" & "_" & crtDate
                                          
             '--NEW SECOND OUTREACH QUERY - PULLS AFTER 4 DAYS--
             DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "qryNeedsSecondOutReach1", _
                    "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\HICS DBs\UPLOADS\" & "Morning_Reports" & "_" & crtDate
             DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "qryNewAssignDateCaseReload", _
                    "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\HICS DBs\UPLOADS\" & "Morning_Reports" & "_" & crtDate
             DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "qryCaseTotalsBySubcategory", _
                    "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\HICS DBs\UPLOADS\" & "Morning_Reports" & "_" & crtDate
             DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "qryHICSNewComment", _
                    "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\HICS DBs\UPLOADS\" & "Morning_Reports" & "_" & crtDate
             DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "qryFraudCases", _
                    "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\HICS DBs\UPLOADS\" & "Morning_Reports" & "_" & crtDate
            
            
            If MsgBox("Go to Upload folder?", vbYesNo) = vbYes Then
                Shell "explorer.exe " & "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\HICS DBs\UPLOADS\", vbNormalFocus
            End If

        Case 4
            DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "qryUnresolved_Cases_No_FollowUp", _
                    "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\HICS DBs\UPLOADS\" & "Non_Response_Cases_To_Close" & "_" & Format(Now(), "yyyymmdd_hhnnss")
                    If MsgBox("Go to Upload folder?", vbYesNo) = vbYes Then
                        Shell "explorer.exe " & "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\HICS DBs\UPLOADS\", vbNormalFocus
                    End If
        End Select
End Sub
Private Sub cmdGenerateEmailBlaster_Click()
    'pulled weekly. cases with only non-response outreachs are pulled and appended to tblHICSEmail. this table is formatted per UPMC Marketing teams specifications
    Dim crtDate As String: crtDate = Format(Now(), "mm_dd_yyyy_hhnnss")
    
    Select Case grpEmailReports.Value
        Case 1
            DoCmd.OpenQuery ("qappHICSEmail")
            DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "qryHICSEmail", _
                "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\HICS DBs\UPLOADS\" & "HICS_EMAIL_BLASTER" & "_" & crtDate
        DoCmd.runSQL ("UPDATE tblHICSEmail SET EXTRACTED = -1, EXTRACT_DATE = '" & crtDate & "' WHERE EXTRACTED = 0;")
        If MsgBox("Go to Upload folder?", vbYesNo) = vbYes Then
            Shell "explorer.exe " & "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\HICS DBs\UPLOADS\", vbNormalFocus
        End If
    End Select
End Sub
Private Sub cmdGetLunchTimes_Click()
    'shows times spent in lunch, when specialist punches in/out from main LandingPage form
    If Nz(txtEventDate.Value, "") <> "" And Nz(txtEventEndDate.Value, "") <> "" Then
        DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12, "qryLunchPunchesByLunchDate", "X:\Member Enrollment\Member Enrollment(Custom)\Marketplace\Toolkits\Employee DB\Lyndsay Reinoso\LunchPunches" & Format(txtEventDate.Value, "yyyymmdd") & "_to_" & Format(txtEventEndDate.Value, "yyyymmdd")
        MsgBox "Export Complete", vbOKOnly
        
    Else: MsgBox "You must supply dates for both fields", vbOKOnly, "DATES ARE BLANK"
    End If
End Sub

Private Sub cmdGetProductivity_Click()

'Get productivty over a date range - located in the left bottom of the admin form
    Select Case grpProductivty
        Case 1
            DoCmd.OpenQuery "qryOutreaches", acViewNormal, acReadOnly
        Case 2
            DoCmd.OpenQuery "qryClosedCases", acViewNormal, acReadOnly
        Case 3
            DoCmd.OpenQuery "qryMonthlyResolvedCases", acViewNormal, acReadOnly
    End Select

End Sub

Private Sub cmdHICSAdHoc_Click()
'Get productivty over a date range - located in the left bottom of the admin form
    Select Case grpHICSAdHoc
        Case 1
            DoCmd.OpenQuery "qryLevel1Total", acViewNormal, acReadOnly
        Case 2
            DoCmd.OpenQuery "qryLevel2Total", acViewNormal, acReadOnly
        Case 3
        
        Case 4
            DoCmd.OpenQuery "qry5DaysFromDue", acViewNormal, acReadOnly
        Case 5
            DoCmd.OpenQuery "qryClosedNonResponse", acViewNormal, acReadOnly
    End Select
End Sub

Private Sub cmdHICSResponse_Click()
'allows user to import the HICS email responses from member so that they are saved to tblComments and can be viewed by specialist
    If IsNull(txtResponseCaseID) Or IsNull(txtEmailResponse) Then
        If MsgBox("Case ID and Email Response cannot be blank", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            Exit Sub
        End If
    End If
    
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("tblComments")
    With rs
        .AddNew
        ![case_id] = txtResponseCaseID.Value
        ![comment] = Replace(txtEmailResponse.Value, vbCrLf, "<br>")
        ![comment_type] = "Member Email Response"
        ![created_date] = Format(Now(), "mm/dd/yyyy hh:mm")
        ![created_by] = ""
        .Update
    End With
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    txtResponseCaseID = Null
    txtEmailResponse = Null
    
    MsgBox ("Item submitted")
End Sub

Private Sub cmdReopenCase_Click()
'middle top of admin form that flips Resolve case flag on tblCaseStatus to "N" and printed flag to "N"
    
    If IsNull(txtCaseID) Or IsNull(cboReopenReason) Then
        If MsgBox("Case ID and Reopen Reason cannot be blank", vbExclamation + vbOKOnly) = vbOK Then Exit Sub
    End If
    
    Dim db As Database: Set db = CurrentDb
    Dim rs As Recordset
    
    Set rs = db.OpenRecordset("tblCaseStatus", dbOpenDynaset)
        rs.FindFirst ("case_id = " & "'" & Me.txtCaseID.Value & "'")
        If rs.NoMatch Then MsgBox ("Case Number not Found; Check case number and try again"): rs.Close: Set rs = Nothing: Exit Sub
        If Not rs.NoMatch Then
            If rs![resolve_case] = "N" Then
            rs.Close: Set rs = Nothing
            MsgBox ("This case has not been resolved and does not need to be reopened")
            
            ElseIf rs![resolve_case] = "Y" Then
                If MsgBox("Case is currently resolved; Do you wish to reopen the case?", vbYesNo) = vbYes Then
                    rs.Edit
                        rs![resolve_case] = "N"
                        rs![printed] = "N"
                        rs![reopened] = 1
                        rs![reopen_reason] = cboReopenReason.Value
                        rs![reopen_date] = Now()
                        rs![scan_uploaded] = 0
                    rs.Update
                    rs.Close
                    
                Set rs = db.OpenRecordset("tblCaseWorkNotes", dbOpenDynaset)
                    rs.FindFirst ("case_id = " & "'" & Me.txtCaseID.Value & "' AND uploaded = 'N' AND is_res = 'Y'")
                    
                    Do While Not rs.NoMatch
                        If rs![prev_res] = 0 Then
                            rs.Delete
                            rs.FindNext ("case_id = " & "'" & Me.txtCaseID.Value & "' AND uploaded = 'N' AND is_res = 'Y'")
                        End If
                    Loop
                    rs.Close: Set rs = Nothing
                    
                    txtCaseID.Value = Null
                    lstCaseSpecialist.Requery
                    cboReopenReason.Value = Null
                    MsgBox ("Case successfully reopened.")
                Else: rs.Close: Set rs = Nothing: MsgBox ("Case not Reopened")
                End If
            End If
        End If
        
        
            
End Sub

Private Sub cmdRunProcess_Click()
'Loads data into database through various queries lsited below
'must have "Microsoft Office 16.0 Object Library" selected to use FileDialog objects

Dim fDialog As FileDialog
Dim filePath As String, fileLocation As String

Select Case grpMornProcesses.Value
    Case 1:
        Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
        fileLocation = "X:\Member Enrollment\Member Enrollment(Custom)\MarketplaceDatabaseProjects\Address Reports\"
        With fDialog
            .AllowMultiSelect = False
            .Title = "Select a file to import"
            .InitialFileName = fileLocation
            .Filters.Clear
            If .Show = 0 Then
                End
            Else
                filePath = .SelectedItems(1)
            End If
        End With

        MsgBox ("This update will take approx. 2 minutes")
        DoCmd.runSQL "DELETE * FROM tblMemberRoster"
        DBEngine.Idle dbRefreshCache
        DoCmd.TransferText acImportDelim, "RosterImportSpec", "tblMemberRoster", filePath, 1
        MsgBox ("Import complete")
    Case 2
        DoCmd.OpenQuery ("qappHICSExtract")
        DoCmd.OpenQuery ("qappCaseStatus")
    Case 3
        DoCmd.OpenQuery ("qappComments")
    Case 4
        DoCmd.OpenQuery ("qappContacts")
    Case Else:
        MsgBox ("No selection made; please make a selection")
End Select
    
End Sub

'Adds an annoucnement to a table that will be displayed ont he main toolkit landingpage
Private Sub cmdSubmitAnnouncement_Click()
Dim rs As Recordset
Dim db As Database
Set db = CurrentDb
Set rs = db.OpenRecordset("tblAnnouncements", dbOpenDynaset)
    rs.AddNew
    rs![Announcement] = txtAnnouncment.Value
    rs![submit_date] = Now()
    rs![End_Date] = DateAdd("d", txtDuration.Value, Now())
    rs.Update
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    txtAnnouncment.Value = ""
    txtDuration.Value = ""
    MsgBox "Announcements Submitted", vbOKOnly, "Successful Submission"

End Sub

Private Sub cmdSubmitEvent_Click()
    Dim rs As Recordset
    Dim db As Database
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblEvents", dbOpenDynaset)
        rs.AddNew
        rs![Event] = txtEvent.Value
        rs![Event_Date] = txtEventDate.Value
        rs![Event_End_Date] = txtEventEndDate.Value
        rs.Update
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    txtEvent.Value = ""
    txtEventDate.Value = ""
    txtEventEndDate.Value = ""
    MsgBox "Event Submitted", vbOKOnly, "Successful Submission"
End Sub


Private Sub grpReports_Click()
    If grpReports.Value = 7 Then
    Me.txtAdHocCaseID.Visible = True
    Else: Me.txtAdHocCaseID.Visible = False
        Me.txtAdHocCaseID.Value = ""
    End If
    
End Sub
'Used to list specialist for a HICS case listed in the Reopen Case section of the admin form
Private Sub txtCaseID_Exit(Cancel As Integer)
    With Me
    txtCaseID.Value = Trim(txtCaseID.Value)
    lstCaseSpecialist.Requery
    End With
End Sub

Private Sub txtCaseIDUpload_Exit(Cancel As Integer)
 With Me
    txtCaseIDUpload.Value = Trim(txtCaseIDUpload.Value)
    End With
End Sub
