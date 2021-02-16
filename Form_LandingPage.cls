VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LandingPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Declare Function apiGetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Dim varArray As Variant
Dim empID As Integer

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Landing page contains a series of buttons that link to other forms within the toolkit or to outside applications
'this also has 2 queries that run on a timer and show user if they have open hics cases and the status of each
'the timer subroutine will close the toolkit if left open over night.
'this also sets permissions their allows user to enter certain sections of toolkit
'-----------------------------------------------------------------------------------------------------------------------------------------------------------

'----BUTTONS----
Private Sub cmd1095A_Click()
    DoCmd.OpenForm "frm1095ASearch", acNormal, , , , acWindowNormal
End Sub

Private Sub cmdAdmin_Click()
    DoCmd.OpenForm "frmAdmin", acNormal, , , , acWindowNormal
End Sub

Private Sub cmdHICS_Click()
    DoCmd.OpenForm "frmHICSUpdate", acNormal, , , , acWindowNormal
End Sub

Private Sub cmdHICSLetters_Click()
    DoCmd.OpenForm "frmMailMerge", acNormal, , , , acWindowNormal
End Sub

Private Sub cmdHICSNotesEdit_Click()
    DoCmd.OpenForm "frmHICSRecordEdit", acNormal, , , , acWindowNormal
End Sub

Private Sub cmdHICSQA_Click()
    DoCmd.OpenForm "frmQA", acNormal, , , , acWindowNormal
End Sub

Private Sub cmdHICSTrends_Click()
    DoCmd.OpenForm "frmHICSTrends", acNormal, , , , acWindowNormal
End Sub

Private Sub cmdHICSUpload_Click()
    DoCmd.OpenForm "frmHICSUpload", acNormal, , , , acWindowNormal
End Sub

Private Sub cmdLunchIn_Click()
Dim rs As Recordset
Dim db As Database
Set db = CurrentDb
MsgBox CDate(Format(Now(), "m/dd/yyyy"))
Set rs = db.OpenRecordset("tblLunchPunches", dbOpenDynaset)
    rs.FindFirst ("[EmployeeID] = " & empID & " AND [LunchDate] = # " & CDate(Format(Now(), "m/dd/yyyy"))) & " #"
    If rs![PunchIn] = "" Or IsNull(rs![PunchIn]) Then
        rs.Edit
        rs![PunchIn] = Format(Now(), "hh:nn:ss")
        rs.Update
    End If
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    MsgBox "Punch Submitted"
    
End Sub
Private Sub cmdLunchOut_Click()
    Dim rs As Recordset
    Dim db As Database
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblLunchPunches", dbOpenDynaset)
    
    rs.AddNew
    rs![employeeId] = empID
    rs![LunchDate] = Format(Now(), "mm/dd/yy")
    rs![PunchOut] = Format(Now(), "hh:nn:ss")
    rs.Update
    
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    
    MsgBox "Punch Submitted"
    
End Sub
Private Sub cmdOpenXWalk_Click()
     DoCmd.OpenForm "frmCrosswalk", acNormal, , , , acWindowNormal
End Sub
Private Sub cmdPAF_Click()
    DoCmd.OpenForm "frmPAF", acNormal, , , , acWindowNormal
End Sub
Private Sub cmdRCNO_Click()
    DoCmd.OpenForm "frmRCNOView", acNormal, , , , acWindowNormal
End Sub
Private Sub cmdSpecialistQA_Click()
    DoCmd.OpenForm "frmSpecialistQA", acNormal, , , acFormAdd, acWindowNormal
End Sub
Private Sub Command14_Click()
    Dim X As Variant
    Dim strProgram As String
    strProgram = "C:\Program Files\IDM Computer Solutions\UltraEdit\uedit64.exe"
    
    X = Shell(strProgram, vbNormalFocus)
End Sub
Private Sub Command260_Click()
    DoCmd.OpenForm "frmHICsSearch", acNormal, , , , acWindowNormal
End Sub
Private Sub Command294_Click()
    DoCmd.OpenForm "frmERRDispute", , , , , , True
End Sub
Private Sub Command295_Click()

End Sub

Private Sub Command309_Click()
    DoCmd.OpenForm "frmERR_Reports", acNormal, , , , acWindowNormal
End Sub

Private Sub Command310_Click()

End Sub


Public Sub Form_Load()

'this sub gets the info of specialist who opened the toolkit and queries the hics tables for open case info, then populates listboxes with results
    Dim strUser As String
    Dim rs As Recordset
    Dim db As Database: Set db = CurrentDb
    Dim strSupport As String
    lstHICSCases.RowSource = ""
    lstHICSComments.RowSource = ""
    
    strUser = fOSUserName()
    Set rs = db.OpenRecordset("Select * From tblEmployeeNames Where NetworkID = " & "'" & strUser & "'", dbOpenSnapshot)
'    rs.MoveLast
'    rs.MoveFirst
    varArray = rs.GetRows(rs.recordCount)
    rs.Close: Set rs = Nothing
    txtAnnouncements.Value = ""
    
        'if specialist is leadership, execute qry for all open cases
        If varArray(0, 0) = 13 Or varArray(0, 0) = 26 Or varArray(0, 0) = 27 Or varArray(0, 0) = 28 Or varArray(0, 0) = 16 Then
            
            Set rs = db.OpenRecordset("SELECT DISTINCT tblCaseStatus.case_id AS Case_ID, tblCaseStatus.pending_helpdesk AS HelpDesk, tblCaseStatus.date_referred AS [Date], tblCaseStatus.Err, tblCaseStatus.err_date AS ERR_Date, tblCMSCaseData.case_level AS Case_Level, tblCMSCaseData.issuer_assign_date AS Assign_Date, tblCMSCaseData.compliance_due_date AS Due_Date, tblEmployeeNames.specialist_name AS Specialist " & _
                " FROM tblCMSCaseData INNER JOIN (tblEmployeeNames INNER JOIN (tblCaseStatus INNER JOIN tblCutlogNotes ON tblCaseStatus.case_id = tblCutlogNotes.case_id) ON tblEmployeeNames.employee_id = tblCutlogNotes.employee_id) ON (tblCMSCaseData.case_id = tblCutlogNotes.case_id) AND (tblCMSCaseData.case_id = tblCaseStatus.case_id) " & _
                " WHERE (((tblCMSCaseData.compliance_due_date)<=Format(DateAdd('d',5,Now()),'mm/dd/yyyy')) AND ((tblCaseStatus.resolve_case)='N') AND ((tblCaseStatus.transferred)=False));", dbOpenSnapshot)
            
            With rs
                If .recordCount > 0 Then
                    .MoveFirst
                    .MoveLast
                    .MoveFirst
                    
                    lstHICSCases.ColumnCount = 9
                    lstHICSCases.Width = "13"
                    lstHICSCases.ColumnWidths = "1 in; 1.0 in; 1.5 in; 0.5 in; 1.5 in; 1 in; 2 in; 1 in; 1 in"
                    Command294.Visible = False
                    
                    lstHICSCases.AddItem .Fields(0).Name & "," & .Fields(1).Name & "," & .Fields(2).Name & "," & .Fields(3).Name & "," & .Fields(4).Name & "," & .Fields(5).Name & _
                         "," & .Fields(6).Name & "," & .Fields(7).Name & "," & .Fields(8).Name
                    Do While Not rs.EOF
                        lstHICSCases.AddItem .Fields(0).Value & "," & .Fields(1).Value & "," & .Fields(2).Value & "," & .Fields(3).Value & "," & .Fields(4).Value & "," & .Fields(5).Value & _
                         "," & .Fields(6).Value & "," & .Fields(7).Value & "," & .Fields(8).Value
                       rs.MoveNext
                    Loop
                    
                End If
            End With
            rs.Close: Set rs = Nothing
        
        'all other specialists qry cases specific to them
        Else
        
            Set rs = db.OpenRecordset("SELECT Distinct tblCaseStatus.case_id AS Case_Number, tblCMSCaseData.compliance_due_date AS Compliance_Date" & _
                                " FROM tblCMSCaseData INNER JOIN (tblEmployeeNames INNER JOIN" & _
                                " (tblCaseStatus INNER JOIN tblCaseWorkNotes ON tblCaseStatus.case_id = tblCaseWorkNotes.case_id)" & _
                                " ON tblEmployeeNames.employee_id = tblCaseWorkNotes.employee_id) ON (tblCMSCaseData.case_id = tblCaseWorkNotes.case_id)" & _
                                " AND (tblCMSCaseData.case_id = tblCaseStatus.case_id)" & _
                                " WHERE (((tblCaseStatus.resolve_case)='N') AND ((tblCaseStatus.transferred)=0) AND ((tblEmployeeNames.NetworkID)=" & "'" & strUser & "'" & "));", dbOpenSnapshot)
            
            With rs
                If .recordCount > 0 Then
                    .MoveFirst
                    .MoveLast
                    .MoveFirst
                    
                    lstHICSCases.ColumnCount = 2
                    lstHICSCases.ColumnWidths = "1.5in;1.5in"
                    Command294.Visible = False
                    
                    lstHICSCases.AddItem .Fields(0).Name & "," & .Fields(1).Name
                    Do While Not rs.EOF
                       lstHICSCases.AddItem .Fields(0).Value & "," & .Fields(1).Value
                       rs.MoveNext
                    Loop
                    
                End If
            End With
            rs.Close: Set rs = Nothing
        
        End If
        
        'query any new casework comments and display to user
        Set rs = db.OpenRecordset("SELECT DISTINCT tblComments.case_id, tblComments.comment, tblComments.comment_type, tblComments.created_date, tblComments.created_by " & _
                                " FROM (tblCaseStatus INNER JOIN tblCutlogNotes ON tblCaseStatus.case_id = tblCutlogNotes.case_id) INNER JOIN tblComments ON tblCaseStatus.case_id = tblComments.case_id " & _
                                " WHERE (((tblCaseStatus.resolve_case)='N') AND ((tblComments.created_date)>=Format(DateAdd('d',-3,Now()),'mm/dd/yyyy')) AND ((tblCutlogNotes.employee_id) =" & varArray(0, 0) & "));", dbOpenSnapshot)
        
        With rs
            If .recordCount > 0 Then
                .MoveFirst
                .MoveLast
                .MoveFirst
                
                lstHICSComments.AddItem .Fields(0).Name & "," & .Fields(1).Name & "," & .Fields(2).Name & "," & .Fields(3).Name & "," & .Fields(4).Name
                Do While Not rs.EOF
                   lstHICSCases.AddItem .Fields(0).Value & "," & .Fields(1).Value & "," & .Fields(2).Value & "," & .Fields(3).Value & "," & .Fields(4).Value
                   rs.MoveNext
                Loop
                
            End If
        End With
        rs.Close: Set rs = Nothing
        db.Close: Set db = Nothing
        
    'set permissions based on values in tblEmployeeNames
    'buttons are enabled/disabled depending upon permissions
    If varArray(5, 0) = True Then
        cmdHICS.Enabled = True
    Else: cmdHICS.Enabled = False
    End If

    If varArray(6, 0) = True Then
        cmdAdmin.Enabled = True
    Else: cmdAdmin.Enabled = False
    End If

    If varArray(7, 0) Then
        cmdHICSQA.Enabled = True
    Else: cmdHICSQA.Enabled = False
    End If

    If varArray(8, 0) Then
        cmdHICSUpload.Enabled = True
    Else: cmdHICSUpload.Enabled = False
    End If

    If varArray(9, 0) Then
        cmdHICSLetters.Enabled = True
    Else: cmdHICSLetters.Enabled = False

    End If

    If varArray(10, 0) Then
        cmdRCNO.Caption = "RCNO"
        cmdRCNO.Enabled = True
    ElseIf Not varArray(10, 0) Then
        cmdRCNO.Caption = "RCNO (READ ONLY)"
        cmdRCNO.Enabled = True
    End If

    If varArray(11, 0) Then
        cmdPAF.Enabled = True
    Else: cmdPAF.Enabled = False
    End If

    If varArray(12, 0) Then
    Else:
    End If

    If varArray(13, 0) Then
    Else:
    End If
    empID = varArray(0, 0)
    Set db = CurrentDb
    Set rs = db.OpenRecordset("Select * From tblAnnouncements Where tblAnnouncements.End_Date >= Format(Now(),'Short Date'); ", dbOpenSnapshot)
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
            While Not rs.EOF
            txtAnnouncements.Value = txtAnnouncements.Value & rs![Announcement] & vbCrLf & vbCrLf
            rs.MoveNext
            Wend
        Else: rs.Close: Set rs = Nothing
        End If
       
        
'        Set rs = db.OpenRecordset("qryDailySupport", dbOpenSnapshot)
'        If rs.RecordCount <= 0 Then
'            rs.Close: Set rs = Nothing
'            Exit Sub
'        End If
'
'        strSupport = rs![Support]
'        lblLeadershipSupport.Caption = "Late Leadership Support: " & strSupport
'        rs.Close: Set rs = Nothing
'        db.Close: Set db = Nothing: Exit Sub
    

End Sub


Function fOSUserName() As String
' Returns the network login name
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

Private Sub Form_Timer()
'form timer fires every 10 mins (600,000 milliseconds)
    If Time() > "1:00:00 AM" And Time() < "2:00:00 AM" Then
        Call TestClose
        Application.CloseCurrentDatabase
    End If

    Form.Refresh
    Form_Load
End Sub

Private Sub Label142_Click()
'Sharepoint - OFF adds and updates link. New Sharepoint site redirects to login screen for Office 365 - normal hyperlink doesn't work
    CreateObject("Shell.application").Open "https://upmchs.sharepoint.com/sites/hpgovernmentprograms/chip/ChipEnrollment/StaffPage/Enrollment%20Services%202011/Lists/Off%20Marketplace%20Adds%20%20Updates/AllItems.aspx?InplviewHashbc115da1-5830-4c30-9051-95fa4b256d82=ShowInGrid=True"
End Sub

Private Sub Label146_Click()
'Sharepoint - self reporting
    CreateObject("Shell.application").Open "https://upmchs.sharepoint.com/sites/hpgovernmentprograms/chip/ChipEnrollment/StaffPage/Enrollment%20Services%202011/Lists/Marketplace%20SelfReporting/AllItems.aspx"
End Sub

Private Sub Label322_Click()
    DoCmd.OpenForm "frmTransperfect", acNormal, , , , acWindowNormal
End Sub

Private Sub Label325_Click()
    DoCmd.OpenForm "frmSysInfo", acNormal, , , , acWindowNormal
End Sub

Private Sub lstHICSCases_DblClick(Cancel As Integer)
    If lstHICSCases.Value <> "" Then
     If varArray(0, 0) = 13 Or varArray(0, 0) = 26 Or varArray(0, 0) = 27 Or varArray(0, 0) = 28 Then
        DoCmd.OpenForm "frmQA", acNormal, , , , acWindowNormal, lstHICSCases.Value
     Else: DoCmd.OpenForm "frmHICSUpdate", acNormal, , , , acWindowNormal, lstHICSCases.Value
    End If
    Else: Exit Sub
    End If
    
    
End Sub
