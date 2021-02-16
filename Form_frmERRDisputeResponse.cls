VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmERRDisputeResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdSearch_Click()
    Dim strSearch As String
    Dim strDisputeType As String
    Dim strIssuerControlID As String
    Dim strDisputeTable As String
    Dim rs As Recordset
    Dim db As Database
    
    cmdSearch.Enabled = False
    strSearch = txtSearchCriteria.Value
    strDisputeType = Nz(DLookup("[Dispute Category]", "tblEnrollmentDisputes", "[ERR Case ID] = " & "'" & strSearch & "'"), "")
    
    Select Case strDisputeType
    Case "Newborn Premium Updates"
        strDisputeTable = "tblNewbornPremiumUpdates"
    Case "Reinstatement End Date 12/31"
        strDisputeTable = "tblReinstatementEndDate"
    Case "HICS Case Only"
        MsgBox "HICS Direct Dispute Case Number"
    Case "Rejected Enrollment"
        strDisputeTable = "tblRejectedEnrollments"
    Case "Enrollment Blocker"
        strDisputeTable = "tblEnrollmentBlocker"
    Case ""
        MsgBox "Dispute not found; check Case ID"
        cmdSearch.Enabled = True
        Exit Sub
    Case Else
        strDisputeTable = "tblDiscrepancyDispute"
    End Select
    
    strSQLSearch = "Select * " & _
    " From [tblEnrollmentDisputes] Inner Join " & strDisputeTable & _
    " ON [tblEnrollmentDisputes].[Issuer Assigned Dispute Control Number] = [" & strDisputeTable & "].[Issuer Assigned Dispute Control Number]" & _
    " WHERE [tblEnrollmentDisputes].[ERR Case ID] = " & "'" & strSearch & "'" & " ;"
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQLSearch)
        
        txtERRCaseID.Value = rs![ERR Case ID]
        txtDisputeCategory.Value = rs![Dispute Category]
        txtBatchNumber.Value = rs![tblEnrollmentDisputes.Batch Number]
        txtExchangeAssignedID.Value = rs![Exchange Assigned Subscriber ID]
        txtDisposition.Value = rs![Disposition]
        txtFileDate.Value = rs![Recent File Name]
        txtCoverageYear.Value = rs![Coverage Year]
        txtInventoryNumber.Value = rs![Internal Inventory Number]
        txtExchangeAssignedPolicyID.Value = rs![Exchange Assigned Policy ID]
        txtDisputeControlNumber.Value = rs![tblEnrollmentDisputes.Issuer Assigned Dispute Control Number]
        txtERR_ReceivedDate.Value = rs![Date Received]
        txtDisputeSpecialist.Value = rs![employee_name]
        txtDispositionDescription.Value = rs![tblEnrollmentDisputes.Description]
        txtInternalNotes.Value = rs![employee_comments]
        txtNotesSubmitDate.Value = rs![submit_date]
        txtNoteSubmitter.Value = rs![employee_name]
        txtCommentsToFFM.Value = rs![Comments]
        If strDisputeTable = "tblDiscrepancyDispute" Then
            txtIssuerValue.Value = rs![Issuer Value]
            txtFFMValue.Value = rs![FFM Value]
            txtHICSCaseID.Value = rs![HICS Case ID]
        End If
        
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    
    cmdSearch.Enabled = True
   
    
    
End Sub

Private Sub cmdSubmitDisposition_Click()
    cmdSubmitDisposition.Enabled = False
    
    Dim rs As Recordset
    Dim db As Database
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblERRDisputeInternalDispositions", dbOpenDynaset)
    If Nz(DLookup("tblERRDisputeInternalDispositions", "ERRCaseID", "[ERRCaseID] = " & "'" & txtERRCaseID.Value & "'"), "") = "" Then
    
    rs.AddNew
        rs![ERRCaseID] = txtERRCaseID.Value
        rs![FFMPolicyID] = txtExchangeAssignedPolicyID.Value
        rs![InternalDisposition] = cmbInternalDisposition.Value
        rs![Employee] = Environ$("username")
        rs![SubmitDate] = Now()
    rs.Update
    Else
        rs.FindFirst "[ERRCaseID] = " & "'" & txtERRCaseID.Value & "'"
        If rs.NoMatch Then
            MsgBox " something has gone horribly wrong"
             cmdSubmitDisposition.Enabled = True
             rs.Close: Set rs = Nothing
             db.Close: Set db = Nothing
             
             Exit Sub
        ElseIf Not rs.NoMatch Then
            rs.Edit
            rs![InternalDispostion] = cmbInternalDispotion.Value
            rs![Employee] = Environ$("username")
            rs![SubmitDate] = Now()
            rs.Update
        End If
    End If
    rs.Close: Set rs = Nothing
    If Nz(txtNewInternalNote.Value, "") <> "" Then
        Set rs = db.OpenRecordset("tblRCNONotes", dbOpenDynaset)
        rs.AddNew
            rs![ERRCaseID] = txtERRCaseID.Value
            rs![InventoryNumber] = txtInventoryNumber.Value
            rs![BatchNumber] = txtBatchNumber.Value
            rs![Note] = txtNewInternalNote.Value
            rs![FFMExchangeAssignedID] = txtExchangeAssignedID.Value
            rs![FFMPolicyID] = txtExchangeAssignedPolicyID.Value
            rs![Employee] = Environ$("username")
            rs![SubmitDate] = Now()
        rs.Update
        rs.Close: Set rs = Nothing
        db.Close: Set db = Nothing
        
    End If
    
    cmdSubmitDisposition.Enabled = True
    
    
    
    
    
End Sub
