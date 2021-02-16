VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmHICSPriorNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim varPrevNotes As Variant
Dim varIntNotes As Variant
Dim varERRNotes As Variant
Dim recordMax As Long, recordCounter As Long
Private Declare Function apiGetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'Purpose of this sub is to allow users to view prior notes on a hics case
'When this form is opened, the case ID from frmHICSUpdate is passed into Form_Load() and seached.
'This sub also has the ability to delete notes on a case
'-----------------------------------------------------------------------------------------------------------------------------------------------------------

Private Sub Form_Open(Cancel As Integer)
    Me.AutoResize = True
    Me.FitToScreen = False
End Sub
Public Sub Form_Load()

On Error GoTo err1:
    'pass case id from main hics form and save in variable
    Dim frm As Form: Set frm = Forms![frmHICSUpdate]
    Dim caseID As String: caseID = frm.txtCaseID
    Dim db As DAO.Database: Set db = CurrentDb
    Dim rs As DAO.Recordset
    
    Set rs = db.OpenRecordset("SELECT * FROM tblCaseWorkNotes WHERE (tblCaseWorkNotes.case_id='" & caseID & "') ORDER BY submit_date", dbOpenSnapshot)
    '################ CASEWORK NOTES #################
    With rs
        If .recordCount > 0 Then
            recordMax = .recordCount - 1
            recordCounter = 0
            varPrevNotes = .GetRows(.recordCount)
        Else
            'if no notes found, display message and close this form
            MsgBox ("No prior casework notes were found.")
            DoCmd.Close acForm, "frmHICSPriorNotes", acSaveNo
            Exit Sub
        End If
    End With
    rs.Close
    
    '################ INTERNAL NOTES #################
    Set rs = db.OpenRecordset("SELECT * FROM tblInternalNotes WHERE case_id = '" & caseID & "';", dbOpenSnapshot)
    With rs
        If .recordCount > 0 Then
            .MoveFirst
            .MoveLast
            .MoveFirst
            varIntNotes = .GetRows(.recordCount)
        End If
    End With
    rs.Close
    
    '################ ERR NOTES #################
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblERRNotes WHERE CaseID = '" & caseID & "' ORDER BY SubmitDate;", dbOpenSnapshot)
    With rs
        If .recordCount > 0 Then
            .MoveFirst
            .MoveLast
            .MoveFirst
            varERRNotes = .GetRows(.recordCount)
        End If
    End With
    
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    
    If Not IsEmpty(varPrevNotes) Then
        Call PopulateNotes(recordCounter)
        Me.txtCaseID = frm.txtCaseID
        Me.txtMemberName = frm.txtFirstName & " " & frm.txtLastName & vbNullString
    End If
    
'error handler
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmHICSPriorNotes; Form_Load()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub
Public Sub PopulateNotes(i As Long)

On Error GoTo err1:

    'CASEWORK NOTES
    txtNoteId.Value = varPrevNotes(15, i)
    txtCaseworkNotes.Value = Replace(varPrevNotes(1, i), "|", vbCrLf & vbCrLf)
    txtSubmitDate.Value = varPrevNotes(2, i)
    txtNoteUploaded.Value = varPrevNotes(4, i)
    txtResolution.Value = varPrevNotes(6, i)
    txtLetterResponse.Value = varPrevNotes(7, i)
    If Not IsNull(varPrevNotes(12, i)) Then: _
        txtSpecialist.Value = DLookup("specialist_name", "tblEmployeeNames", "employee_id = " & CInt(varPrevNotes(12, i)))
    If Not IsNull(varPrevNotes(13, i)) Then: _
        txtInteractionType.Value = DLookup("contact_type", "tblHICSInteractionType", "ID = " & CInt(varPrevNotes(13, i)))
    If Not IsNull(varPrevNotes(14, i)) Then: _
        txtCaseAction.Value = DLookup("action", "tblHICSCaseAction", "ID = " & CInt(varPrevNotes(14, i)))
    txtRecordCounter.Value = i + 1
    txtRecordMax.Value = recordMax + 1
    
    'FIND MATCHING INTERNAL NOTE
    If Not IsEmpty(varIntNotes) Then
        For j = 0 To UBound(varIntNotes, 2)
            If varIntNotes(5, j) = varPrevNotes(15, i) Then
                txtInternalNotes.Value = varIntNotes(2, j)
                txtInternalNoteId.Value = varIntNotes(0, j)
                Exit For
            Else
                txtInternalNotes.Value = Null
                txtInternalNoteId.Value = Null
            End If
        Next j
    End If
    
    'FIND MATCHING ERR NOTE
    If Not IsEmpty(varERRNotes) Then
        For j = 0 To UBound(varERRNotes, 2)
            If varERRNotes(8, j) = varPrevNotes(15, i) Then
                txtERRNote.Value = varERRNotes(2, j)
                txtERRNoteId.Value = varERRNotes(0, j)
                Exit For
            Else
                txtERRNote.Value = Null
                txtERRNoteId.Value = Null
            End If
        Next j
    End If
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmHICSPriorNotes; PopulateNotes()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub

Private Sub btnClose_Click()
    DoCmd.Close acForm, "frmHICSPriorNotes", acSaveNo
End Sub
Private Sub btnSubmit_Click()
    Dim frm As Form
    Set frm = Forms![frmHICSUpdate]
    frm.SubmitCaseWorkNote
    MsgBox ("Item submitted - nice job, kid.")
    DoCmd.Close acForm, "frmHICSPriorNotes", acSaveNo
End Sub

'DELETE CASEWORK NOTE
Private Sub cmdDelete_Click()
    
    Dim empID As Integer: empID = DLookup("employee_id", "tblEmployeeNames", "NetworkID = '" & fOSUserName() & "'")
    
    'VALIDATION - user can only delete their own notes that haven't been uploaded to CMS yet
    If varPrevNotes(12, recordCounter) <> empID Then
        If MsgBox("Cannot edit a record submitted by another specialist.", vbExclamation + vbOKOnly, "Error") = vbOK Then: Exit Sub
    End If
    If txtNoteUploaded.Value = "Y" Then
        If MsgBox("Cannot edit record - casework note has already been loaded to CMS system.", vbExclamation + vbOKOnly, "Error") = vbOK Then: Exit Sub
    End If
    
On Error GoTo err1:
    Dim resolution As String: resolution = varPrevNotes(6, recordCounter)
    Dim caseID As String: caseID = varPrevNotes(0, recordCounter)
    Dim recordId As Long: recordId = varPrevNotes(15, recordCounter)
    Dim searchStr As String: searchStr = "case_id = '" & varPrevNotes(0, recordCounter) & "' AND ID = " & varPrevNotes(15, recordCounter) & " AND uploaded = 'N'"
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
                    
                    'reload main form to show case is reopened
                    Call Forms("frmHICSUpdate").cmdResetForm_Click
                    Forms![frmHICSUpdate].txtCaseID.Value = caseID
                    Call Forms("frmHICSUpdate").cmdSearchHICSCase_Click
                    DoCmd.Close acForm, "frmHICSPriorNotes", acSaveNo
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
                Call Form_Load
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
            Call LogError(Err.Number & " " & Err.Description, "frmHICSPriorNotes; cmdDelete_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
End Sub
'EDIT PREV CASEWORK NOTE
Private Sub cmdEdit_Click()

    If varPrevNotes(6, recordCounter) <> fOSUserName() Then
        MsgBox ("Cannot edit a record submitted by another specialist.")
        Exit Sub
    End If
    If txtNoteUploaded.Value = "Y" Then
        MsgBox ("Cannot edit record - casework note has already been loaded to CMS system.")
        Exit Sub
    End If

    btnSubmit.Visible = True
    
    
End Sub

Private Sub cmdPrevNote_Click()
    'view previous note
    If recordMax = 0 Then Exit Sub
    If recordCounter = 0 Then
        Exit Sub
    Else
        recordCounter = recordCounter - 1
        txtRecordCounter.Value = recordCounter + 1
        Call PopulateNotes(recordCounter)
    End If
End Sub
Private Sub cmdNextNote_Click()
    'view next note
    If recordMax = 0 Then Exit Sub
    If recordCounter = recordMax Then
        Exit Sub
    Else
        recordCounter = recordCounter + 1
        txtRecordCounter.Value = recordCounter + 1
        Call PopulateNotes(recordCounter)
    End If
End Sub



Function fOSUserName() As String
' Returns the network id
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
