VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmHICsSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'form tied RCNO form

'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'This allows user to search for HICS case either by member id or by first/last name
'If results are found, user can view more info by clicking Command22 btn and view results in frmQA
'-----------------------------------------------------------------------------------------------------------------------------------------------------------


'search, sets rowsource of listbox to results
Private Sub cmdSearch_Click()


    Select Case optSearchType.Value
        Case 1 'search by member id
            If txtMMBRSearch.Value <> "" Then
                If Len(txtMMBRSearch.Value) > 9 Then
                    txtMMBRSearch.Value = Left(txtMMBRSearch.Value, 9)
                End If
                
                lstResults.RowSource = "SELECT tblMemberDemo.member_id, tblMemberDemo.case_id, tblCMSCaseData.issuer_assign_date " & _
                                        "FROM tblCMSCaseData INNER JOIN tblMemberDemo ON tblCMSCaseData.case_id = tblMemberDemo.case_id " & _
                                        "WHERE Left([tblMemberDemo].[member_id],9) = '" & txtMMBRSearch.Value & "';"
        
                Form.Refresh
                
            Else
                MsgBox "Search Boxes cannot be blank", vbOKOnly
                lstResults.RowSource = ""
            End If
           
        Case 2 'search by name
            Dim strCriteria As String
            
            If IsNull(txtFirstName.Value) And IsNull(txtLastName) Then
                MsgBox "Search Boxes cannot be blank", vbOKOnly
                lstResults.RowSource = ""
                Exit Sub
            ElseIf Not IsNull(txtFirstName) And Not IsNull(txtLastName) Then
                strCriteria = " WHERE tblContact.first_name LIKE '" & txtFirstName & "' AND tblContact.last_name LIKE '" & txtLastName & "'"
            ElseIf Not IsNull(txtFirstName) And IsNull(txtLastName) Then
                strCriteria = " WHERE tblContact.first_name LIKE '" & txtFirstName & "'"
            ElseIf IsNull(txtFirstName) And Not IsNull(txtLastName) Then
                strCriteria = " WHERE tblContact.last_name LIKE '" & txtLastName & "'"
            End If
                 
               lstResults.RowSource = "SELECT member_id, tblCMSCaseData.case_id, tblContact.first_name, tblContact.last_name, Format(tblCMSCaseData.issuer_assign_date, 'mm/dd/yyyy') AS [issuer_assign_date] " & _
                                    " FROM (tblCMSCaseData INNER JOIN tblContact ON tblCMSCaseData.case_id = tblContact.case_id) LEFT JOIN tblMemberDemo ON tblCMSCaseData.case_id = tblMemberDemo.case_id " & strCriteria
                Form.Refresh
    End Select
   
End Sub

'passes selected case from lstBox to frmQA
Private Sub Command22_Click()

    If lstResults.ItemsSelected.Count > 0 Then
        If CurrentProject.AllForms("frmQA").IsLoaded = True Then
            DoCmd.Close acForm, "frmQA", acSaveNo
            DoCmd.OpenForm "frmQA", acNormal, , , , acWindowNormal, lstResults.Value
        Else
            DoCmd.OpenForm "frmQA", acNormal, , , , acWindowNormal, lstResults.Value
        End If
    Else
        MsgBox "Cannot open - please select a record from the listbox to open.", vbOKOnly
    End If
    
    
End Sub

'if member number is being passed on form open
Private Sub Form_Open(Cancel As Integer)
    optSearchType_Click
    Dim strRefSBSNum As String
     lstResults.RowSource = ""
       
    If Not IsNull(Forms!frmHICsSearch.OpenArgs) Then
        strRefSBSNum = Forms!frmHICsSearch.OpenArgs
        If Len(strRefSBSNum) > 0 Then
            txtMMBRSearch.Value = strRefSBSNum
            cmdSearch_Click
        End If
    Else: lstResults.RowSource = ""
    End If
End Sub

Private Sub optSearchType_Click()
clearBoxes

Select Case optSearchType.Value
    Case 1
        txtMMBRSearch.Visible = True
    Case 2
        txtFirstName.Visible = True
        txtLastName.Visible = True
End Select


End Sub

Public Sub clearBoxes()

txtMMBRSearch.Value = Null
txtFirstName.Value = Null
txtLastName.Value = Null
txtMMBRSearch.Visible = False
txtFirstName.Visible = False
txtLastName.Visible = False


End Sub
