VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmRCNOView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Declare Function apiGetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public avarRecords As Variant
Public currentIndex As Long
Public rcrdNum As Long

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



Private Sub cboDisputedItem_AfterUpdate()
    If cboDisputedItem.Value = cboDisputedItem.ItemData(13) Then
        txtIssuerValue.Visible = True
    Else: txtIssuerValue.Visible = False
    End If
    
    
End Sub

Private Sub cboFormType_AfterUpdate()
If cboFormType.Value = cboFormType.ItemData(0) Then
    cboDisputedItem.Visible = True
Else: cboDisputedItem.Visible = False
End If
If cboFormType.Value = cboFormType.ItemData(2) Then
    txtAPTC.Visible = True
    txtTotalPremium.Visible = True
    txtIssuerValue.Visible = True
Else: txtAPTC.Visible = False: txtTotalPremium.Visible = False: txtIssuerValue.Visible = False
If cboFormType.Value = cboFormType.ItemData(6) Then
    grpAgentBroker.Visible = True
Else: grpAgentBroker.Visible = False
End If

End If


    
    

End Sub

Private Sub cmd1095Dispute_Click()
    Dim db As Database
    Dim rs As Recordset
    
On Error GoTo err1:
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tbl1095disputes", dbOpenDynaset)
    rs.AddNew
    
    If txt35.Value = "" Or IsNull(txt35.Value) Then
        rs![member_number] = txt34.Value
    Else: rs![member_number] = txt35.Value
    End If
    
    rs![exchange_subscriber_number] = txt25.Value
    rs![exchange_member_number] = txt28.Value
     
     If txt8.Value = "" Or IsNull(txt8.Value) Then
        rs![member_last_name] = txt7.Value
    Else: rs![member_last_name] = txt8.Value
    End If
    
    If txt2.Value = "" Or IsNull(txt2.Value) Then
        rs![member_first_name] = txt1.Value
    Else: rs![member_first_name] = txt2.Value
    End If
    
    If txt11.Value = "" Or IsNull(txt11.Value) Then
        rs![date_of_birth] = txt10.Value
    Else: rs![date_of_birth] = txt11.Value
    End If
    
    If txt17.Value = "" Or IsNull(txt17.Value) Then
        rs![social_security_number] = txt16.Value
    Else: rs![social_security_number] = txt17.Value
    End If
    
    rs![employee_name] = fOSUserName()
    rs![submit_date] = Now()
    rs.Update
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    
MsgBox "1095 Dispute Submitted"

err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmRCNOView; cmd1095Dispute_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select

End Sub

Private Sub cmdFindDisputes_Click()
    DoCmd.OpenForm "frmViewDisputes", acNormal, , , acFormReadOnly
End Sub


Private Sub cmdFindDisputes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    txt25.BackColor = RGB(225, 225, 0)
'    txt.BackColor = RGB(255, 255, 255)
    
End Sub

Private Sub cmdFindHICS_Click()
'txt32 is Issuer Subscriber ID - Issuer Value
    DoCmd.OpenForm "frmHICsSearch", acNormal, , , , acWindowNormal, Me.txt32.Value
    
End Sub
Private Sub btnViewNotes_Click()

'opens form to view internal notes
'txt28 is FFM Assigned ID
    DoCmd.OpenForm "frmRCNOviewNotes", acNormal, , , , acWindowNormal, Me.txt25.Value
End Sub

Private Sub NoteCheck()
'FFM Assigned ID - when value in box changes, query if notes exist

txtNotes.Value = Null
If DSum("[ID]", "[tblRCNONotes]", "[FFMExchangeAssignedID] = '" & txt25.Value & "'") > 0 Then txtNotes.Value = vbCrLf & vbCrLf & vbCrLf & "     VIEW PREVIOUS NOTES"

End Sub
Private Sub cmdFrwd_Click()
    Dim strSearch As String: strSearch = txtSearchInv.Value
    
    
    If rcrdNum <= currentIndex + 1 Then
        MsgBox "No more records"
    ElseIf rcrdNum > currentIndex + 1 Then
        ClearForm
    txtSearchInv.Value = strSearch
        currentIndex = currentIndex + 1
        For i = -1 To 164
            Controls("txt" & i + 1).Value = avarRecords(i + 7, currentIndex)
           If Controls("txt" & i + 1).Value = "I" Or Controls("txt" & i + 1).Value = "L" Then
                Controls("txt" & i + 1).BackColor = RGB(240, 240, 0)
                Controls("txt" & i).BackColor = RGB(240, 240, 0)
                Controls("txt" & i - 1).BackColor = RGB(240, 240, 0)
            
            ElseIf Controls("txt" & i + 1).Value = "U" Then
                Controls("txt" & i + 1).BackColor = RGB(240, 0, 0)
                Controls("txt" & i).BackColor = RGB(240, 0, 0)
                Controls("txt" & i - 1).BackColor = RGB(240, 0, 0)
            Else: Controls("txt" & i + 1).BackColor = RGB(255, 255, 255)
            End If
             
            Select Case Controls("txt" & i + 1).Value
            Case "M"
                Controls("txt" & i + 1).ControlTipText = "Exact match: No Action"
            Case "I"
                Controls("txt" & i + 1).ControlTipText = "Non-Match: Issuer to update to FFM value"
            Case "F"
                Controls("txt" & i + 1).ControlTipText = "Non-Match: FFM to update to Issuer value"
            Case "L"
                Controls("txt" & i + 1).ControlTipText = "Non-Match: FFM and Issuer to update to value listed in the FFM field"
            Case "U"
                Controls("txt" & i + 1).ControlTipText = "Uprocessable: Issuer to correct file"
            Case "C"
                Controls("txt" & i + 1).ControlTipText = "Match due to Change in Circumstances or maintencance transaction, no action"
            Case "D"
                Controls("txt" & i + 1).ControlTipText = "Did not compare: No Action"
            Case "G"
                Controls("txt" & i + 1).ControlTipText = "Non-match: FFM to update based on ER&R determination"
            Case "J"
                Controls("txt" & i + 1).ControlTipText = "Non-Match: Issuer to update based on ER&R determination"
            Case "K"
                Controls("txt" & i + 1).ControlTipText = "Non-Match: both FFM and Issuer to update to value listed in the FFM field"
            End Select
            If (i + 1 = 13) Or (i + 1 = 14) Or (i + 1 = 4) Or (i + 1 = 5) Then
            Controls("txt" & i + 1).ControlTipText = ""
            End If
        Next i
        
        FormatSSN txt16: FormatSSN txt17
        
        FormatDate txt0
        FormatDate txt10: FormatDate txt11
        FormatDate txt88: FormatDate txt89
        FormatDate txt91: FormatDate txt92
        FormatDate txt97: FormatDate txt98
        FormatDate txt100: FormatDate txt101
        FormatDate txt106: FormatDate txt107
        FormatDate txt109: FormatDate txt110
        FormatDate txt115: FormatDate txt116
        FormatDate txt118: FormatDate txt119
        FormatDate txt124: FormatDate txt125
        FormatDate txt127: FormatDate txt128
        
        FormatZip txt55: FormatZip txt56
        FormatZip txt70: FormatZip txt71
        NoteCheck
        
        txt0.BackColor = RGB(59, 129, 148)
        OverallFlagDescription
        
    End If
    
End Sub

Private Sub cmdIssuerUpdated_Click()

Dim db As Database: Set db = CurrentDb
Dim rs As Recordset

On Error GoTo err1:
    Set rs = db.OpenRecordset("tblIssuerUpdatedRecords", dbOpenDynaset)
        rs.AddNew
        rs![Inventory_Number] = txt134.Value
        rs![Batch_Number] = txt136.Value
        rs![member_number] = txt35.Value
        rs![Notes] = txtNotes.Value
        rs![employee_name] = fOSUserName()
        rs.Update
        rs.Close: Set rs = Nothing
        db.Close: Set db = Nothing
        MsgBox "Added Record to Table"
        txtNotes.Value = ""
        
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmRCNOView; cmdIssuerUpdated_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
End Sub

Private Sub cmdPrev_Click()
    Dim strSearch As String: strSearch = txtSearchInv.Value
     
     If currentIndex - 1 < 0 Then
        MsgBox "No more records"
    ElseIf currentIndex - 1 >= 0 Then
        ClearForm
        txtSearchInv.Value = strSearch
        currentIndex = currentIndex - 1
        For i = -1 To 164
         Controls("txt" & i + 1).Value = avarRecords(i + 7, currentIndex)
          If Controls("txt" & i + 1).Value = "I" Or Controls("txt" & i + 1).Value = "L" Then
                Controls("txt" & i + 1).BackColor = RGB(240, 240, 0)
                Controls("txt" & i).BackColor = RGB(240, 240, 0)
                Controls("txt" & i - 1).BackColor = RGB(240, 240, 0)
            
            ElseIf Controls("txt" & i + 1).Value = "U" Then
                Controls("txt" & i + 1).BackColor = RGB(240, 0, 0)
                Controls("txt" & i).BackColor = RGB(240, 0, 0)
                Controls("txt" & i - 1).BackColor = RGB(240, 0, 0)
            Else: Controls("txt" & i + 1).BackColor = RGB(255, 255, 255)
            End If
            Select Case Controls("txt" & i + 1).Value
            Case "M"
                Controls("txt" & i + 1).ControlTipText = "Exact match: No Action"
            Case "I"
                Controls("txt" & i + 1).ControlTipText = "Non-Match: Issuer to update to FFM value"
            Case "F"
                Controls("txt" & i + 1).ControlTipText = "Non-Match: FFM to update to Issuer value"
            Case "L"
                Controls("txt" & i + 1).ControlTipText = "Non-Match: FFM and Issuer to update to value listed in the FFM field"
            Case "U"
                Controls("txt" & i + 1).ControlTipText = "Uprocessable: Issuer to correct file"
            Case "C"
                Controls("txt" & i + 1).ControlTipText = "Match due to Change in Circumstances or maintencance transaction, no action"
            Case "D"
                Controls("txt" & i + 1).ControlTipText = "Did not compare: No Action"
            Case "G"
                Controls("txt" & i + 1).ControlTipText = "Non-match: FFM to update based on ER&R determination"
            Case "J"
                Controls("txt" & i + 1).ControlTipText = "Non-Match: Issuer to update based on ER&R determination"
            Case "K"
                Controls("txt" & i + 1).ControlTipText = "Non-Match: both FFM and Issuer to update to value listed in the FFM field"
            End Select
            If (i + 1 = 13) Or (i + 1 = 14) Or (i + 1 = 4) Or (i + 1 = 5) Then
            Controls("txt" & i + 1).ControlTipText = ""
            End If

        Next i

        FormatSSN txt16: FormatSSN txt17
        
        FormatDate txt0
        FormatDate txt10: FormatDate txt11
        FormatDate txt88: FormatDate txt89
        FormatDate txt91: FormatDate txt92
        FormatDate txt97: FormatDate txt98
        FormatDate txt100: FormatDate txt101
        FormatDate txt106: FormatDate txt107
        FormatDate txt109: FormatDate txt110
        FormatDate txt115: FormatDate txt116
        FormatDate txt118: FormatDate txt119
        FormatDate txt124: FormatDate txt125
        FormatDate txt127: FormatDate txt128



        FormatZip txt55: FormatZip txt56
        FormatZip txt70: FormatZip txt71
        NoteCheck
        
        txt0.BackColor = RGB(59, 129, 148)
        OverallFlagDescription
    End If
End Sub

Private Sub cmdSubmitNote_Click()
    Dim db As Database
    Dim rs As Recordset
    
On Error GoTo err1:
    Set db = CurrentDb
    cmdSubmitNote.Enabled = False
        If Nz(txtNotes.Value, "") <> "" Then
        Set rs = db.OpenRecordset("tblRCNONotes", dbOpenDynaset)
        rs.AddNew
        rs![InventoryNumber] = txt134.Value
        rs![BatchNumber] = txt136.Value
        rs![Note] = txtNotes.Value
        rs![FFMPolicyID] = IIf(Nz(txt37.Value, "") = "", txt38.Value, txt37.Value)
        rs![FFMExchangeAssignedID] = IIf(Nz(txt25.Value, "") = "", txt26.Value, txt25.Value)
        rs![Employee] = fOSUserName()
        rs![SubmitDate] = Now()
        rs.Update
        rs.Close: Set rs = Nothing
        MsgBox "Note Submitted", vbOKOnly, "NOTE SUBMITTED"
        Else: MsgBox "Please fill out the 'Notes' box to submit notes", vbExclamation
        End If
        
    cmdSubmitNote.Enabled = True
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmRCNOView; cmdSubmitNote_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
    
End Sub

'search button
'txtbox number is field number - 7

Private Sub Command10_Click()
    Dim rs As Recordset
    Dim db As Database
    Dim strSearch As String:
    If txtSearchInv.Value <> "" Or Not IsNull(txtSearchInv.Value) Then
        strSearch = txtSearchInv.Value
    Else: MsgBox "Seach field cannot be blank": Exit Sub
    End If

On Error GoTo err1:
    Dim strYear As String
    Set db = CurrentDb
    currentIndex = 0
    ClearForm
    txtSearchInv.Value = strSearch
    
    'CASE STATEMENT - EACH CASE VALUE BELOW CORRESPONDS TO A RADIO BUTTON IN
    'THE 'BATCH' OPTION GROUP ON tblRCNOView
    Select Case grpYear.Value
        Case 1: strYear = "tblRCNO"
        Case 2: strYear = "tblRCNO18"
        Case 3: strYear = "tblRCNO1809"
        Case 4: strYear = "tblRCNO1808"
        Case 5: strYear = "tblRCNO1807"
        Case 6: strYear = "tblRCNO1806"
        Case 7: strYear = "tblRCNO1805"
        Case 8: strYear = "tblRCNO9"
        Case 9: strYear = "tblRCNOTrain"
        Case 10: strYear = "tblRCNO1810"
        Case 11: strYear = "tblRCNO1811"
        Case 12: strYear = "tblRCNO1812"
        Case 13: strYear = "tblRCNO1813"
        Case 14: strYear = "tblRCNO1901"
        Case 15: strYear = "tblRCNO1814"
        Case 16: strYear = "tblRCNO1902"
        Case 17: strYear = "tblRCNO1903"
        Case 18: strYear = "tblRCNO1904"
        Case 19: strYear = "tblRCNO1815"
        Case 20: strYear = "tblRCNO1905"
        Case 21: strYear = "tblRCNO1906"
        Case 22: strYear = "tblRCNO16"
        Case 23: strYear = "tblRCNO1907"
        Case 24: strYear = "tblRCNO20"
        
        
    Case Else: MsgBox "You must select a Year": Exit Sub
    End Select
    
    If grpSearch.Value = 1 Then
        'Set rs = db.OpenRecordset("SELECT " & strYear & ".*, tblRCNO1804AppID.[Application ID], tblRCNO1804AppID.[Insurance Application Origin Type] FROM " & strYear & " LEFT JOIN tblRCNO1804AppID ON " & strYear & ".[FFM Internal Inventory Record Number] = tblRCNO1804AppID.[FFM Internal Inventory Record Number] WHERE tblRCNO18.[FFM Internal Inventory Record Number] = " & txtSearchInv.Value, dbOpenSnapshot)
        Set rs = db.OpenRecordset("SELECT * From " & strYear & " WHERE [FFM Internal Inventory Record Number] = " & txtSearchInv.Value, dbOpenSnapshot)
        cmdFrwd.Enabled = False
        cmdPrev.Enabled = False
    
    ElseIf grpSearch.Value = 2 Then
        Set rs = db.OpenRecordset( _
        "SELECT * From " & strYear & " WHERE [ISSUER BENE SBS EXCHG ASG ID] = " & "'" & txtSearchInv.Value & "'" & "OR [FFM BENE SBS EXCHG ASG ID] = " & "'" & txtSearchInv.Value & "'" & " OR [ISSUER BENE EXCHG ASGNED ID] = " & "'" & txtSearchInv.Value & "'" & "OR [FFM BENE EXCHG ASGNED ID] = " & "'" & txtSearchInv.Value & "'" & "Order By [ISSUER BENE ISSR ASGNED ID], [ISSUER PLAN BNFT STRT DT]", dbOpenSnapshot)
'            "SELECT * From " & strYear & " WHERE [ISSUER BENE SBS EXCHG ASG ID] = " & "'" & txtSearchInv.Value & "'" & "OR [FFM BENE SBS EXCHG ASG ID] = " & "'" & txtSearchInv.Value & "'" & "Order By [ISSUER BENE ISSR ASGNED ID], [ISSUER PLAN BNFT STRT DT]", dbOpenSnapshot)
        cmdFrwd.Enabled = True
        cmdPrev.Enabled = True
    
    ElseIf grpSearch.Value = 3 Then
        Set rs = db.OpenRecordset( _
        "SELECT * From " & strYear & " WHERE [ISSUER BENE EXCHG ASGNED ID] = " & "'" & txtSearchInv.Value & "'" & "OR [FFM BENE EXCHG ASGNED ID] = " & "'" & txtSearchInv.Value & "'" & "Order By [ISSUER BENE ISSR ASGNED ID], [ISSUER PLAN BNFT STRT DT]", dbOpenSnapshot)
        cmdFrwd.Enabled = True
        cmdPrev.Enabled = True
    
    ElseIf grpSearch.Value = 4 Then
        Set rs = db.OpenRecordset( _
        "SELECT * From " & strYear & " WHERE [ISSUER BENE SBS ISSR ASG ID] = " & "'" & txtSearchInv.Value & "'" & "OR [FFM BENE SBS ISSR ASG ID] = " & "'" & txtSearchInv.Value & "'" & "Order By [ISSUER BENE ISSR ASGNED ID], [ISSUER PLAN BNFT STRT DT]", dbOpenSnapshot)
        cmdFrwd.Enabled = True
        cmdPrev.Enabled = True
    ElseIf grpSearch.Value = 5 Then
        Set rs = db.OpenRecordset( _
        "SELECT * From " & strYear & " WHERE [ISSUER BENE ISSR ASGNED ID] = " & "'" & txtSearchInv.Value & "'" & "OR [FFM BENE ISSR ASGNED ID] = " & "'" & txtSearchInv.Value & "'" & "Order By [ISSUER BENE ISSR ASGNED ID], [ISSUER PLAN BNFT STRT DT]", dbOpenSnapshot)
        cmdFrwd.Enabled = True
        cmdPrev.Enabled = True
     ElseIf grpSearch.Value = 6 Then
        Set rs = db.OpenRecordset( _
        "SELECT * From " & strYear & " WHERE [ISS PLAN EXCHG ASG PLCY NUM] = " & "'" & txtSearchInv.Value & "'" & "OR [FFM PLAN EXCHG ASG PLCY NUM] = " & "'" & txtSearchInv.Value & "'" & "Order By [ISSUER BENE ISSR ASGNED ID], [ISSUER PLAN BNFT STRT DT]", dbOpenSnapshot)
        cmdFrwd.Enabled = True
        cmdPrev.Enabled = True
    ElseIf grpSearch.Value = 7 Then
        Set rs = db.OpenRecordset( _
        "SELECT * From " & strYear & " WHERE [TrainingCode] = " & "'" & txtSearchInv.Value & "'" & " Order By [ISSUER BENE ISSR ASGNED ID], [ISSUER PLAN BNFT STRT DT]", dbOpenSnapshot)
        cmdFrwd.Enabled = True
        cmdPrev.Enabled = True
    End If
'    ElseIf grpSearchBy.Value = 2 Then
'        If grpSearch.Value = 1 Then
'            Set rs = db.OpenRecordset("SELECT * From tblRCNO WHERE [FFM Internal Inventory Record Number] = " & txtSearchInv.Value, dbOpenSnapshot)
'            cmdFrwd.Enabled = False
'            cmdPrev.Enabled = False
'
'        ElseIf grpSearch.Value = 2 Then
'            Set rs = db.OpenRecordset( _
'            "SELECT * From tblRCNO WHERE [FFM BENE SBS EXCHG ASG ID] = " & "'" & txtSearchInv.Value & "'", dbOpenSnapshot)
'            cmdFrwd.Enabled = True
'            cmdPrev.Enabled = True
'
'        ElseIf grpSearch.Value = 3 Then
'            Set rs = db.OpenRecordset( _
'            "SELECT * From tblRCNO WHERE [FFM BENE EXCHG ASGNED ID] = " & "'" & txtSearchInv.Value & "'", dbOpenSnapshot)
'            cmdFrwd.Enabled = True
'            cmdPrev.Enabled = True
'
'        ElseIf grpSearch.Value = 4 Then
'            Set rs = db.OpenRecordset( _
'            "SELECT * From tblRCNO WHERE [FFM BENE SBS ISSR ASG ID] = " & "'" & txtSearchInv.Value & "'", dbOpenSnapshot)
'            cmdFrwd.Enabled = True
'            cmdPrev.Enabled = True
'        ElseIf grpSearch.Value = 5 Then
'            Set rs = db.OpenRecordset( _
'            "SELECT * From tblRCNO WHERE [FFM BENE ISSR ASGNED ID] = " & "'" & txtSearchInv.Value & "'", dbOpenSnapshot)
'            cmdFrwd.Enabled = True
'            cmdPrev.Enabled = True
'        End If

    If rs.recordCount > 0 Then
        rs.MoveLast
        rs.MoveFirst
        rcrdNum = rs.recordCount
        avarRecords = rs.GetRows(rs.recordCount)
      
    Else: MsgBox ("No records"): Exit Sub
    End If
        
    'GET COLUMN COUNT. OLDER RCNO VERSION HAVE 172 FIELDS WHILE NEWER HAS 181
    Dim tbl As DAO.TableDef
    Dim colCount As Integer
    Set tbl = db.TableDefs(strYear)
    colCount = tbl.Fields.Count - 8
    Set tbl = Nothing
        
    'LOOP THROUGH avarRecords AND ASSIGN TO CONTROLS ON FORM
    For i = -1 To colCount
        Controls("txt" & i + 1).Value = avarRecords(i + 7, currentIndex)
           If Controls("txt" & i + 1).Value = "I" Or Controls("txt" & i + 1).Value = "L" Then
                Controls("txt" & i + 1).BackColor = RGB(240, 240, 0)
                Controls("txt" & i).BackColor = RGB(240, 240, 0)
                Controls("txt" & i - 1).BackColor = RGB(240, 240, 0)
            
            ElseIf Controls("txt" & i + 1).Value = "U" Then
                Controls("txt" & i + 1).BackColor = RGB(240, 0, 0)
                Controls("txt" & i).BackColor = RGB(240, 0, 0)
                Controls("txt" & i - 1).BackColor = RGB(240, 0, 0)
            Else: Controls("txt" & i + 1).BackColor = RGB(255, 255, 255)
            End If

            Select Case Controls("txt" & i + 1).Value
            Case "M"
                Controls("txt" & i + 1).ControlTipText = "Exact match: No Action"
            Case "I"
                Controls("txt" & i + 1).ControlTipText = "Non-Match: Issuer to update to FFM value"
            Case "F"
                Controls("txt" & i + 1).ControlTipText = "Non-Match: FFM to update to Issuer value"
            Case "L"
                Controls("txt" & i + 1).ControlTipText = "Non-Match: FFM and Issuer to update to value listed in the FFM field"
            Case "U"
                Controls("txt" & i + 1).ControlTipText = "Uprocessable: Issuer to correct file"
            Case "C"
                Controls("txt" & i + 1).ControlTipText = "Match due to Change in Circumstances or maintencance transaction, no action"
            Case "D"
                Controls("txt" & i + 1).ControlTipText = "Did not compare: No Action"
            Case "G"
                Controls("txt" & i + 1).ControlTipText = "Non-match: FFM to update based on ER&R determination"
            Case "J"
                Controls("txt" & i + 1).ControlTipText = "Non-Match: Issuer to update based on ER&R determination"
            Case "K"
                Controls("txt" & i + 1).ControlTipText = "Non-Match: both FFM and Issuer to update to value listed in the FFM field"
            
            End Select
            If (i + 1 = 13) Or (i + 1 = 14) Or (i + 1 = 4) Or (i + 1 = 5) Then
            Controls("txt" & i + 1).ControlTipText = ""
            End If
    Next i
    
    'FORMAT DATES, SSN, ZIP CODE
    FormatSSN txt16: FormatSSN txt17
    FormatDate txt0
    FormatDate txt10: FormatDate txt11
    FormatDate txt88: FormatDate txt89
    FormatDate txt91: FormatDate txt92
    FormatDate txt97: FormatDate txt98
    FormatDate txt100: FormatDate txt101
    FormatDate txt106: FormatDate txt107
    FormatDate txt109: FormatDate txt110
    FormatDate txt115: FormatDate txt116
    FormatDate txt118: FormatDate txt119
    FormatDate txt124: FormatDate txt125
    FormatDate txt127: FormatDate txt128
    
    FormatZip txt55: FormatZip txt56
    FormatZip txt70: FormatZip txt71
    NoteCheck
    
    txt0.BackColor = RGB(59, 129, 148)
    OverallFlagDescription
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    If txt73.Value <> "" Then
            txt73.ControlTipText = DLookup("County", "tblFFMCountyCodes", "[FIPS State county code] = " & "'" & txt73.Value & "'")
    End If
    If txt74.Value <> "" Then
        txt74.ControlTipText = Nz(DLookup("CountyDescription", "tblMC400CountyCodes", "CountyCode = " & "'" & txt74.Value & "'"), "")
        
    End If
    If txt22.Value <> "" Then
        txt22.ControlTipText = DLookup("Description", "tblRelCodes", "[ASC X12 Value] = " & "'" & txt22.Value & "'")
    End If
    If txt23.Value <> "" Then
        txt23.ControlTipText = DLookup("Description", "tblRelCodes", "[ASC X12 Value] = " & "'" & txt23.Value & "'")
    End If
    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmRCNOView; Command10_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub

Sub FormatSSN(ctrl As Control)
    If ctrl.Value <> "" And Not IsNull(ctrl.Value) Then
        ctrl.Value = Left(ctrl.Value, 3) & "-" & Mid(ctrl.Value, 4, 2) & "-" & Right(ctrl.Value, 4)
    Else: Exit Sub
    End If

End Sub

Sub FormatDate(ctrl As Control)
    If ctrl.Value <> "" And Not IsNull(ctrl.Value) Then
    ctrl.Value = Mid(ctrl.Value, 5, 2) & "/" & Right(ctrl.Value, 2) & "/" & Left(ctrl.Value, 4)
    Else: Exit Sub
    End If
    
End Sub

Sub FormatZip(ctrl As Control)
    If Len(ctrl.Value) = 9 Then
        ctrl.Value = Left(ctrl.Value, 5) & "-" & Right(ctrl.Value, 4)
    Else: Exit Sub
    End If
    
End Sub
Function FormatDateUpload(ctrl As Control) As String
    
    FormatDateUpload = Right(ctrl.Value, 4) & Left(ctrl.Value, 2) & Mid(ctrl.Value, 4, 2)
    
    
End Function

Private Sub btnSubmitDispute_Click()
Dim db As Database
Dim rs As Recordset
Set db = CurrentDb

    
    'VALIDATION
    If CheckSubmission(cboFormType.ListIndex) = True Then
        If MsgBox("Error submitting form. See 'Submission Errors' for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then
            txtErrorBar.SetFocus
            Exit Sub
        End If
    End If
    
    
On Error GoTo err1:

    Select Case cboFormType.Value
        Case cboFormType.ItemData(0) 'Discrepancy Dispute
            Set rs = db.OpenRecordset("tblDiscrepancyDispute", dbOpenDynaset)
            rs.FindFirst ("[FFM Internal Inventory Record Number] = " & "'" & txt134.Value & "'" & "And [Description] = " & "'" & cboDisputedItem.Value & "'")
                If Not rs.NoMatch Then
                        MsgBox "Unable to dispute because of duplication"
                        rs.Close: Set rs = Nothing
                        Exit Sub
                    End If
            rs.AddNew
            'rs![Issuer Value] = txtIssuerValue.Value
            rs![Benefit Start Date] = FormatDateUpload(txt88)
            rs![Benefit End Date] = FormatDateUpload(txt91)
            rs![Comments] = Clean(txtComments.Value)
            rs![HICS Case ID] = txtHICSCase.Value
            Select Case cboDisputedItem.Value
                 Case cboDisputedItem.ItemData(0) 'Agent/Broker Name
                     MsgBox "Agent/Broker Information"
                     rs![Description] = cboDisputedItem.ItemData(0)
                     rs![FFM Value] = txt151.Value & " " & txt154.Value & " " & txt157.Value & " " & txt160.Value
                 
                 Case cboDisputedItem.ItemData(1) 'Agent/Broker NPN
                     MsgBox "Agent/Broker NPN"
                     rs![Description] = cboDisputedItem.ItemData(1)
                     rs![FFM Value] = txt148.Value
                 
                 Case cboDisputedItem.ItemData(2) 'Applied APTC Amount
                     MsgBox "Applied APTC Amount"
                     rs![Description] = cboDisputedItem.ItemData(2)
                     rs![FFM Value] = txt94.Value
                     rs![Issuer Value] = txt95.Value
                 Case cboDisputedItem.ItemData(3) 'Benefit Start Date
                     MsgBox "Benefit Start Date"
                     rs![Description] = cboDisputedItem.ItemData(3)
                     rs![FFM Value] = FormatDateUpload(txt88)
                     rs![Issuer Value] = FormatDateUpload(txt89)
                 Case cboDisputedItem.ItemData(4) 'CSR Amount
                     MsgBox "CSR Amount"
                     rs![Description] = cboDisputedItem.ItemData(4)
                     rs![FFM Value] = txt103.Value
                     rs![Issuer Value] = txt104.Value
                 Case cboDisputedItem.ItemData(5) 'End of Year Termination Indicator
                     MsgBox "Dispute Not Submitted": Exit Sub
                     'rs![Description] = cboDisputedItem.ItemData(5)
                     'rs![FFM Value] =
                 
                 Case cboDisputedItem.ItemData(6) 'Initial Premium Paid Status
                     MsgBox "Initial Premium Paid Status Indicator"
                     rs![Description] = cboDisputedItem.ItemData(6)
                     rs![FFM Value] = txt130.Value
                     rs![Issuer Value] = txt131.Value
                 Case cboDisputedItem.ItemData(7) 'Issuer Assigned Member ID
                     MsgBox "Issuer Assigned Member ID"
                     rs![Description] = cboDisputedItem.ItemData(7)
                     rs![FFM Value] = txt34.Value
                     rs![Issuer Value] = txt35.Value
                 Case cboDisputedItem.ItemData(8) 'Issuer Assigned Policy ID
                     MsgBox "Issuer Assigned Policy ID"
                     rs![Description] = cboDisputedItem.ItemData(8)
                     rs![FFM Value] = txt40.Value
                     rs![Issuer Value] = txt41.Value
                 Case cboDisputedItem.ItemData(9) 'Issuer Assigned Subscriber ID
                     MsgBox "Issuer Assigned Subscriber ID"
                     rs![Description] = cboDisputedItem.ItemData(9)
                     rs![FFM Value] = txt31.Value
                     rs![Issuer Value] = txt32.Value
                 Case cboDisputedItem.ItemData(10) 'Issuer End Date Earlier Than FFM
                     MsgBox "Issur End Date Earlier Than FFM"
                     rs![Description] = cboDisputedItem.ItemData(10)
                     rs![FFM Value] = FormatDateUpload(txt91)
                     rs![Issuer Value] = FormatDateUpload(txt92)
                 Case cboDisputedItem.ItemData(11) 'Issuer End Date Later Than FFM
                     MsgBox "Issuer End Date Later Than FFM"
                     rs![Description] = cboDisputedItem.ItemData(11)
                     rs![FFM Value] = FormatDateUpload(txt91)
                     rs![Issuer Value] = FormatDateUpload(txt92)
                 Case cboDisputedItem.ItemData(12) 'DO NOT USE ---Paid Through Date
                     MsgBox ""
                     rs![Description] = cboDisputedItem.ItemData(12)
                     rs![FFM Value] = FormatDateUpload(txt91)
                     rs![Issuer Value] = FormatDateUpload(txt92.Value)
                 Case cboDisputedItem.ItemData(13) 'Prior Year - End Date
                     MsgBox "Prior Year - End Date"
                    rs![Description] = cboDisputedItem.ItemData(13)
                    rs![FFM Value] = FormatDateUpload(txt91)
                    If txtIssuerValue.Value = "" Or IsNull(txtIssuerValue.Value) Then
                        rs![Issuer Value] = FormatDateUpload(txt92)
                    Else: rs![Issuer Value] = txtIssuerValue.Value
                    End If
                 Case cboDisputedItem.ItemData(14) 'QHP ID
                     MsgBox "QHP ID"
                     rs![Description] = cboDisputedItem.ItemData(14)
                     rs![FFM Value] = txt85.Value
                     rs![Issuer Value] = txt86.Value
                 Case cboDisputedItem.ItemData(15) 'Tobacco Status
                     MsgBox "Tobbacco Status"
                     rs![Description] = cboDisputedItem.ItemData(15)
                     rs![FFM Value] = txt82.Value
                     rs![Issuer Value] = txt83.Value
                 Case cboDisputedItem.ItemData(16) 'Total Premium Amount
                     MsgBox "Total Premium Amount"
                     rs![Description] = cboDisputedItem.ItemData(16)
                     rs![FFM Value] = txt112.Value
                     rs![Issuer Value] = txt113.Value
                Case cboDisputedItem.ItemData(17) 'Gender
                     MsgBox "Gender"
                     rs![Description] = cboDisputedItem.ItemData(17)
                     rs![FFM Value] = txt13.Value
                     rs![Issuer Value] = txt14.Value
                    
                Case Else: MsgBox "You must select a dispute type": rs.Close: Set rs = Nothing: Exit Sub
            End Select
        
        Case cboFormType.ItemData(1) 'Enrollment Blocker
            Set rs = db.OpenRecordset("tblEnrollmentBlocker", dbOpenDynaset)
            rs.AddNew
            rs![QHP Variant ID] = Right(txt85.Value, 2)
            rs![Issuer Start Date] = FormatDateUpload(txt89)
            rs![FFM Start Date] = FormatDateUpload(txt88)
            rs![FFM End Date] = FormatDateUpload(txt91)
            rs![Total Premium] = txt113.Value
            rs![APTC Amount] = txt95.Value
            rs![CSR] = txt104.Value
            rs![HICS Case ID] = txtHICSCase.Value
            rs![Comments] = Clean(txtComments.Value)
        Case cboFormType.ItemData(2) 'Newborn Premium Updates
            Set rs = db.OpenRecordset("tblNewbornPremiumUpdates", dbOpenDynaset)
            rs.AddNew
            rs![Issuer Value] = txtIssuerValue.Value
            rs![Subscriber Total Premium] = txtTotalPremium.Value
            rs![Subscriber APTC] = txtAPTC.Value
            
        Case cboFormType.ItemData(3) 'Reinstatement End Date
            Set rs = db.OpenRecordset("tblReinstatementEndDate", dbOpenDynaset)
            rs.AddNew
            rs![FFM Benefit Start Date] = FormatDateUpload(txt88)
            rs![FFM Benefit End Date] = FormatDateUpload(txt91)
            
        Case cboFormType.ItemData(4) 'Rejected Enrollments
            Set rs = db.OpenRecordset("tblRejectedEnrollments", dbOpenDynaset)
            rs.AddNew
        Case cboFormType.ItemData(5) 'Mailing Address Dispute
            Set rs = db.OpenRecordset("tblMailingAddressDispute", dbOpenDynaset)
            rs.AddNew
            rs![Issuer Mail Address Line 1] = txt59.Value
            rs![Issuer Mail Address Line 2] = txt62.Value
            rs![Issuer Mail Address City] = txt50.Value
            rs![Issuer Mail Address State] = txt68.Value
            rs![Issuer Mail Address ZIP] = txt71.Value
            MsgBox "Mailing Address Dispute Submitted"
        Case cboFormType.ItemData(6) 'Agent Broker Dispute
            Set rs = db.OpenRecordset("tblAgentBrokerDisputes", dbOpenDynaset)
            rs.AddNew
            Select Case grpAgentBroker
            Case 1
            rs![Action] = "Agent Broker Name/NPN - Remove"
            Case 2
            rs![Action] = "Agent Broker Name/NPN - Update"
            rs![Issuer Agent Broker First Name] = txt149.Value
            rs![Issuer Agent Broker Middle Name] = txt152.Value
            rs![Issuer Agent Broker Last Name] = txt155.Value
            rs![Issuer Agent Broker Suffix] = txt158.Value
            rs![Issuer Agent Broker NPN] = txt161.Value
            
            End Select
            
        Case Else: MsgBox "Please select a Dispute Type": Exit Sub
    End Select
    
            rs![Batch Number] = txt136.Value
            rs![FFM Internal Inventory Record Number] = txt134.Value
            rs![Exchange Assigned Policy ID] = txt37.Value
            rs![Exchange Assigned Subscriber ID] = txt25.Value
            rs![CoverageYear] = txt139.Value
            rs![employee_comments] = Clean(txtNotes.Value & vbNullString)
            rs![employee_name] = fOSUserName()
            rs![Extracted] = "N"
            rs![Issuer Assigned Dispute Control Number] = txt134.Value & txt136.Value & Format(Now(), "nnss") & fOSUserName()
            rs![submit_date] = Now()
            rs.Update
            rs.Close: Set rs = Nothing
            db.Close: Set db = Nothing
    
    txtNotes.Value = ""
    txtComments.Value = ""
    txtIssuerValue.Value = ""
    txtHICSCase.Value = ""
    txtTotalPremium.Value = ""
    txtAPTC.Value = ""
    cboDisputedItem.Value = ""
    cboFormType.Value = ""
   
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmRCNOView; btnSubmitDispute_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
   
End Sub

Private Sub Command380_Click()
    Dim db As Database
    Dim rs As Recordset
    
On Error GoTo err1:
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblRCNI_Issues", dbOpenDynaset)
    rs.AddNew
    rs![FFM Inventory Number] = txt134.Value
    rs![Batch_Number] = txt136.Value
    rs![Notes] = txtNotes.Value
    rs![employee_name] = fOSUserName()
    rs.Update
    rs.Close: Set rs = Nothing
    db.Close: Set db = Nothing
    
    txtNotes.Value = ""
    MsgBox "Sent to RCNI Issues Form"
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "frmRCNOView; Command380_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
End Sub




Private Sub Form_Load()
    Dim ctrl As Control
        For Each ctrl In Controls
            If TypeOf ctrl Is TextBox Then
                If ctrl.Name <> "txtErrorBar" And ctrl.Name <> "txtErrorBox" Then
                    ctrl.BackColor = RGB(255, 255, 255)
                End If
            End If
        Next ctrl
        
    If Not IsNull(Forms!frmRCNOView.OpenArgs) Then
    strRefSubscriberID = Forms!frmRCNOView.OpenArgs
        If Len(strRefSubscriberID) > 0 Then
            strRefSubscriberID = Left(strRefSubscriberID, 9)
            txtSearchInv.Value = strRefSubscriberID
        End If
    End If
End Sub
Private Sub OverallFlagDescription()
 Select Case txt133.Value
    Case "M"
        txtRecordDescription.Value = _
        "1:1 Record Match and all fields that are compared match between the issuer and FFM record; " & _
        "may contain field-level discrepancies flagged ‘D’ – Did Not Compare"
    
    Case "E"
        txtRecordDescription.Value = _
        "Uneven Match with Issuer Action Required to resolve; " & _
        "additionally, one or more field-level discrepancies flagged for Issuer update and one or more fields flagged for FFM update; " & _
        "however, due to uneven matching of records within the enrollment group, only limited FFM updates may be applied in this cycle."
    
    Case "N"
        txtRecordDescription.Value = _
        "1:1 Match with Issuer Action Required to resolve; at least one field flagged for Issuer action; one or more fields may also be flagged for FFM update."
    
    Case "P"
        txtRecordDescription.Value = _
        "Uneven Match with Issuer Action Required to resolve; " & _
        "no field-level discrepancies flagged for Issuer update but one or more fields flagged for FFM update; " & _
        "however due to uneven matching of records within the enrollment group, only limited FFM updates may be applied in this cycle."
    
    Case "Z"
        txtRecordDescription.Value = _
        "1:1 Match with no Issuer Action required; no field-level discrepancies flagged for Issuer update, but one or more fields flagged for FFM update."
    
    Case "C"
        txtRecordDescription.Value = _
        "Cancelation - The enrollment record has been canceled in the Issuer’s System or the FFM"
    
    Case "B"
        txtRecordDescription.Value = _
        "FFM Generated CIC Enrollment. The member is part of an FFM Enrollment CIC Generated condition and is used to complete the entire Enrollment Group to facilitate matching the Issuer version of a CIC to the FFM version of that CIC. No Issuer action required."
    
    Case "U"
        txtRecordDescription.Value = _
        "Unprocessable Record with Issuer Action Required to resolve: Significant data issues have caused removal of the record from the matching process; " & _
        "such issues include, but are not limited to, poorly formed date or financial values, or dates in the incorrect plan year."
    
    Case "F"
        txtRecordDescription.Value = _
        "Unmatched FFM Enrollment: No matching Issuer record found for this FFM record, by"
    
    Case "G"
        txtRecordDescription.Value = _
        "'Leftover' FFM Enrollment: After determining the best match for all records in the Quattro, FFM records remain that cannot be matched to an Issuer record."
    
    Case "I"
        txtRecordDescription.Value = _
        "Unmatched Issuer Enrollment: No matching FFM record found for this Issuer record,"
    
    Case "R"
        txtRecordDescription.Value = _
        "'Leftover' Issuer Enrollment: After determining the best match for all records in the Quattro, Issuer records remain that cannot be matched to an FFM record."
    
    Case "W"
        txtRecordDescription.Value = _
        "Non-Match – Many to Many: The best match cannot be found for multiple records with the same Quattro Key on both the FFM and Issuer side, all records remain unmatched"
    
    Case "L"
        txtRecordDescription.Value = _
        "Non-Match – Many to One/One to Many: The best match cannot be found for multiple records on either the FFM or Issuer side to one record on the other side with the same Quattro Key, all records remain unmatched"
    
    Case "D"
        txtRecordDescription.Value = _
        "Duplicate Issuer Record (all duplicates flagged with D)"
        
    Case "V"
        txtRecordDescription.Value = _
        "Coverage Span to be Cancelled on FFM; record is active on FFM but exists outside the bounds of active coverage submitted by the Issuer on the RCNI file Note: Although no further Issuer action is required, Issuers are encouraged to review V records to ensure the cancellation of the coverage span is appropriate and not due to a technical problem with the Issuer’s extract"
    Case "S"
        txtRecordDescription.Value = _
        "FFM Data Artifact Due to Financial Eligibility Change (No Issuer Action Required); record is associated with another matched record on the file, but due to a financial eligibility change not conveyed to the Issuer, more spans exist on the FFM side; coverage represented on the S record(s) will be considered in conjunction with any matched records for the same Exchange-Assigned Policy ID when applying updates to FFM"
    End Select
End Sub

Sub ClearForm()
    Dim ctrl As Control
    For Each ctrl In Controls
        If TypeOf ctrl Is TextBox Then
            ctrl.Value = ""
        End If
    Next
    txtErrorBox.BackColor = RGB(166, 166, 166)
    txtErrorBar.BackColor = RGB(166, 166, 166)
    
        
End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim boolPrem As Boolean
    boolPrem = Nz(Forms!frmRCNOView.OpenArgs, -1)
    Enabler (boolPrem)
    
    
End Sub
Function Clean(txtClean As String)
        txtClean = Replace(txtClean, vbLf, " ")
        txtClean = Replace(txtClean, vbTab, " ")
        txtClean = Replace(txtClean, vbCr, " ")
        txtClean = Replace(txtClean, vbCrLf, " ")
        txtClean = Replace(txtClean, vbNewLine, " ")
        txtClean = Replace(txtClean, ";", ":")
        txtClean = Replace(txtClean, "<", "(")
        txtClean = Replace(txtClean, ">", ")")
        txtClean = Replace(txtClean, Chr(160), " ")
        txtClean = Replace(txtClean, "'", " ")
        txtClean = Replace(txtClean, Chr(146), " ")
        txtClean = Replace(txtClean, Chr(39), " ")
        Trim (txtClean)
        Clean = txtClean
End Function



Private Sub txtSearchInv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Command10.SetFocus
        Command10_Click
        
    End If
    
End Sub

Private Sub Enabler(boolState As Boolean)
    
    cboFormType.Enabled = boolState
    cboDisputedItem.Enabled = boolState
    txtIssuerValue.Enabled = boolState
    txtHICSCase.Enabled = boolState
    Command380.Enabled = boolState
    txtComments.Enabled = boolState
    txtTotalPremium.Enabled = boolState
    txtAPTC.Enabled = boolState
    txtNotes.Enabled = boolState
    
    btnSubmitDispute.Enabled = boolState
    cmdIssuerUpdated.Enabled = boolState
    cmd1095Dispute.Enabled = boolState
    
End Sub
Public Function CheckSubmission(submitType As Integer) As Boolean

    Dim msgStr As String, headerStr As String, footerStr As String, resultStr As String
    Dim regex As RegExp: Set regex = New RegExp
    headerStr = "<ul>"
    footerStr = "</ul>"
    
    Select Case submitType
        Case 0 'Discrepancy Dispute
            With regex
                .MultiLine = False
                .Global = True
                .IgnoreCase = True
                .Pattern = "^[a-zA-Z0-9]{11}$"
            End With
            If txtHICSCase <> "" Then
                If regex.test(txtHICSCase.Value) = False Then
                    msgStr = msgStr & "<li><b>HICS Case ID</b> is invalid. Please correct and resubmit.</li>"
                End If
            End If

            Select Case cboDisputedItem.Value
                 Case cboDisputedItem.ItemData(0) 'Agent/Broker Name
                    If txt151.Value = "" Or IsNull(txt151.Value) Then
                        msgStr = msgStr & "<li><b>Broker First Name - FFM Value</b> cannot be blank.</li>"
                    End If
                    If txt157.Value = "" Or IsNull(txt157.Value) Then
                        msgStr = msgStr & "<li><b>Broker Last Name - FFM Value</b> cannot be blank.</li>"
                    End If
                 Case cboDisputedItem.ItemData(1) 'Agent/Broker NPN
                    If txt148.Value = "" Or IsNull(txt148.Value) Then
                        msgStr = msgStr & "<li><b>Broker NPN - FFM Value</b> cannot be blank.</li>"
                    End If
                 Case cboDisputedItem.ItemData(2) 'Applied APTC Amount
                    If txt94.Value = "" Or IsNull(txt94.Value) Then
                        msgStr = msgStr & "<li><b>APTC Amount - FFM Value</b> cannot be blank.</li>"
                    End If
                    If txt95.Value = "" Or IsNull(txt95.Value) Then
                        msgStr = msgStr & "<li><b>APTC Amount - Issuer Value</b> cannot be blank.</li>"
                    End If
                 Case cboDisputedItem.ItemData(3) 'Benefit Start Date
                    If txt88.Value = "" Or IsNull(txt88.Value) Then
                        msgStr = msgStr & "<li><b>Benefit Start Date - FFM Value</b> cannot be blank.</li>"
                    End If
                    If txt89.Value = "" Or IsNull(txt89.Value) Then
                        msgStr = msgStr & "<li><b>Benefit Start Date - Issuer Value</b> cannot be blank.</li>"
                    End If
                 Case cboDisputedItem.ItemData(4) 'CSR Amount
                    If txt103.Value = "" Or IsNull(txt103.Value) Then
                        msgStr = msgStr & "<li><b>CSR Amount - FFM Value</b> cannot be blank.</li>"
                    End If
                    If txt104.Value = "" Or IsNull(txt104.Value) Then
                        msgStr = msgStr & "<li><b>CSR Amount - Issuer Value</b> cannot be blank.</li>"
                    End If
                 Case cboDisputedItem.ItemData(5) 'End of Year Termination Indicator
                 
                 Case cboDisputedItem.ItemData(6) 'Initial Premium Paid Status
                    If txt130.Value = "" Or IsNull(txt130.Value) Then
                        msgStr = msgStr & "<li><b>Premium Paid Indicator - FFM Value</b> cannot be blank.</li>"
                    End If
                    If txt131.Value = "" Or IsNull(txt131.Value) Then
                        msgStr = msgStr & "<li><b>Premium Paid Indicator - Issuer Value</b> cannot be blank.</li>"
                    End If
                 Case cboDisputedItem.ItemData(7) 'Issuer Assigned Member ID
                    If txt34.Value = "" Or IsNull(txt34.Value) Then
                        msgStr = msgStr & "<li><b>Issuer Member ID - FFM Value</b> cannot be blank.</li>"
                    End If
                    If txt35.Value = "" Or IsNull(txt35.Value) Then
                        msgStr = msgStr & "<li><b>Issuer Member ID - Issuer Value</b> cannot be blank.</li>"
                    End If
                 Case cboDisputedItem.ItemData(8) 'Issuer Assigned Policy ID
                    If txt40.Value = "" Or IsNull(txt40.Value) Then
                        msgStr = msgStr & "<li><b>Issuer Assigned Policy ID - FFM Value</b> cannot be blank.</li>"
                    End If
                    If txt41.Value = "" Or IsNull(txt41.Value) Then
                        msgStr = msgStr & "<li><b>Issuer Assigned Policy ID - Issuer Value</b> cannot be blank.</li>"
                    End If
                 Case cboDisputedItem.ItemData(9) 'Issuer Assigned Subscriber ID
                    If txt31.Value = "" Or IsNull(txt31.Value) Then
                        msgStr = msgStr & "<li><b>Issuer Subscriber ID - FFM Value</b> cannot be blank.</li>"
                    End If
                    If txt32.Value = "" Or IsNull(txt32.Value) Then
                        msgStr = msgStr & "<li><b>Issuer Subscriber ID - Issuer Value</b> cannot be blank.</li>"
                    End If
                 Case cboDisputedItem.ItemData(10) 'Issuer End Date Earlier Than FFM
                    If txt91.Value = "" Or IsNull(txt91.Value) Then
                        msgStr = msgStr & "<li><b>Benefit End Date - FFM Value</b> cannot be blank.</li>"
                    End If
                    If txt92.Value = "" Or IsNull(txt92.Value) Then
                        msgStr = msgStr & "<li><b>Benefit End Date - Issuer Value</b> cannot be blank.</li>"
                    End If
                 Case cboDisputedItem.ItemData(11) 'Issuer End Date Later Than FFM
                    If txt91.Value = "" Or IsNull(txt91.Value) Then
                        msgStr = msgStr & "<li><b>Benefit End Date - FFM Value</b> cannot be blank.</li>"
                    End If
                    If txt92.Value = "" Or IsNull(txt92.Value) Then
                        msgStr = msgStr & "<li><b>Benefit End Date - Issuer Value</b> cannot be blank.</li>"
                    End If
                 Case cboDisputedItem.ItemData(12) 'DO NOT USE ---Paid Through Date
                    If txt91.Value = "" Or IsNull(txt91.Value) Then
                        msgStr = msgStr & "<li><b>Benefit End Date - FFM Value</b> cannot be blank.</li>"
                    End If
                    If txt92.Value = "" Or IsNull(txt92.Value) Then
                        msgStr = msgStr & "<li><b>Benefit End Date - Issuer Value</b> cannot be blank.</li>"
                    End If
                 Case cboDisputedItem.ItemData(13) 'Prior Year - End Date
                    If txt91.Value = "" Or IsNull(txt91.Value) Then
                        msgStr = msgStr & "<li><b>Benefit End Date - FFM Value</b> cannot be blank.</li>"
                    End If
                    If txtIssuerValue.Value = "" Or IsNull(txtIssuerValue.Value) Then
                        If txt92.Value = "" Or IsNull(txt92.Value) Then
                            msgStr = msgStr & "<li><b>Benefit End Date - Issuer Value</b> cannot be blank.</li>"
                        End If
                    End If
                 Case cboDisputedItem.ItemData(14) 'QHP ID
                    If txt85.Value = "" Or IsNull(txt85.Value) Then
                        msgStr = msgStr & "<li><b>QHPID - FFM Value</b> cannot be blank.</li>"
                    End If
                    If txt86.Value = "" Or IsNull(txt86.Value) Then
                        msgStr = msgStr & "<li><b>QHPID - Issuer Value</b> cannot be blank.</li>"
                    End If
                 Case cboDisputedItem.ItemData(15) 'Tobacco Status
                    If txt82.Value = "" Or IsNull(txt82.Value) Then
                        msgStr = msgStr & "<li><b>Tobacco Indicator - FFM Value</b> cannot be blank.</li>"
                    End If
                    If txt83.Value = "" Or IsNull(txt83.Value) Then
                        msgStr = msgStr & "<li><b>Tobacco Indicator - Issuer Value</b> cannot be blank.</li>"
                    End If
                 Case cboDisputedItem.ItemData(16) 'Total Premium Amount
                    If txt112.Value = "" Or IsNull(txt112.Value) Then
                        msgStr = msgStr & "<li><b>Total Prem Amount - FFM Value</b> cannot be blank.</li>"
                    End If
                    If txt113.Value = "" Or IsNull(txt113.Value) Then
                        msgStr = msgStr & "<li><b>Total Prem Amount - Issuer Value</b> cannot be blank.</li>"
                    End If
                Case cboDisputedItem.ItemData(17) 'Gender
                    If txt13.Value = "" Or IsNull(txt13.Value) Then
                        msgStr = msgStr & "<li><b>Gender - FFM Value</b> cannot be blank.</li>"
                    End If
                    If txt14.Value = "" Or IsNull(txt14.Value) Then
                        msgStr = msgStr & "<li><b>Gender - Issuer Value</b> cannot be blank.</li>"
                    End If
                Case Else:
                    If cboDisputedItem.Value = "" Or IsNull(cboDisputedItem.Value) Then
                        msgStr = msgStr & "<li>You must select a <b>Dispute Type</b>.</li>"
                    End If
            End Select

        Case 1 'Enrollment Blocker
            With regex
                .MultiLine = False
                .Global = True
                .IgnoreCase = True
                .Pattern = "^[a-zA-Z0-9]{11}$"
            End With
            If txtHICSCase.Value = "" Or IsNull(txtHICSCase.Value) Then
                msgStr = msgStr & "<li><b>HICS Case ID</b> cannot be blank.</li>"
            End If
            If Not IsNull(txtHICSCase) Then
                If regex.test(txtHICSCase.Value) = False Then
                    msgStr = msgStr & "<li><b>HICS Case ID</b> is invalid. Please correct and resubmit.</li>"
                End If
            End If
            If txt85.Value = "" Or IsNull(txt85.Value) Then
                msgStr = msgStr & "<li><b>QHPID - FFM Value</b> cannot be blank.</li>"
            End If
            If txt89.Value = "" Or IsNull(txt89.Value) Then
                msgStr = msgStr & "<li><b>Benefit Start Date - Issuer Value</b> cannot be blank.</li>"
            End If
            If txt88.Value = "" Or IsNull(txt88.Value) Then
                msgStr = msgStr & "<li><b>Benefit Start Date - FFM Value</b> cannot be blank.</li>"
            End If
            If txt91.Value = "" Or IsNull(txt91.Value) Then
                msgStr = msgStr & "<li><b>Benefit End Date - FFM Value</b> cannot be blank.</li>"
            End If
            If txt113.Value = "" Or IsNull(txt113.Value) Then
                msgStr = msgStr & "<li><b>Total Prem Amount - Issuer Value</b> cannot be blank.</li>"
            End If
            If txt95.Value = "" Or IsNull(txt95.Value) Then
                msgStr = msgStr & "<li><b>APTC Amount - Issuer Value</b> cannot be blank.</li>"
            End If
            If txt104.Value = "" Or IsNull(txt104.Value) Then
                msgStr = msgStr & "<li><b>CSR Amount - Issuer Value</b> cannot be blank.</li>"
            End If
            If txtComments.Value = "" Or IsNull(txtComments.Value) Then
                msgStr = msgStr & "<li><b>Notes</b> cannot be blank.</li>"
            End If
        Case 2 'Newborn Premium Updates
            If txtIssuerValue.Value = "" Or IsNull(txtIssuerValue.Value) Then
                msgStr = msgStr & "<li><b>Issuer Value</b> cannot be blank.</li>"
            End If
            If txtTotalPremium.Value = "" Or IsNull(txtTotalPremium.Value) Then
                msgStr = msgStr & "<li><b>Total Premium</b> cannot be blank.</li>"
            End If
            If txtAPTC.Value = "" Or IsNull(txtAPTC.Value) Then
                msgStr = msgStr & "<li><b>APTC</b> cannot be blank.</li>"
            End If
        Case 3 'Reinstatement End Date
            If txt88.Value = "" Or IsNull(txt88.Value) Then
                msgStr = msgStr & "<li><b>Benefit Start Date - FFM Value</b> cannot be blank.</li>"
            End If
            If txt91.Value = "" Or IsNull(txt91.Value) Then
                msgStr = msgStr & "<li><b>Benefit End Date - FFM Value</b> cannot be blank.</li>"
            End If
        Case 4 'Rejected Enrollments
        
        Case 5 'Mailing Address Dispute
            If txt59.Value = "" Or IsNull(txt59.Value) Then
                msgStr = msgStr & "<li><b>Street Address - Issuer Value</b> cannot be blank.</li>"
            End If
            If txt62.Value = "" Or IsNull(txt62.Value) Then
                msgStr = msgStr & "<li><b>Street Address 2 - Issuer Value</b> cannot be blank.</li>"
            End If
            If txt50.Value = "" Or IsNull(txt50.Value) Then
                msgStr = msgStr & "<li><b>City - Issuer Value</b> cannot be blank.</li>"
            End If
            If txt68.Value = "" Or IsNull(txt68.Value) Then
                msgStr = msgStr & "<li><b>State - Issuer Value</b> cannot be blank.</li>"
            End If
            If txt71.Value = "" Or IsNull(txt71.Value) Then
                msgStr = msgStr & "<li><b>Zip - Issuer Value</b> cannot be blank.</li>"
            End If
        Case 6 'Agent Broker Dispute
            Select Case grpAgentBroker
            Case 2
            If txt149.Value = "" Or IsNull(txt149.Value) Then
                msgStr = msgStr & "<li><b>Broker NPN - Issuer Value</b> cannot be blank.</li>"
            End If
            If txt152.Value = "" Or IsNull(txt152.Value) Then
                msgStr = msgStr & "<li><b>Broker First Name - Issuer Value</b> cannot be blank.</li>"
            End If
            If txt158.Value = "" Or IsNull(txt158.Value) Then
                msgStr = msgStr & "<li><b>Broker Last Name - Issuer Value</b> cannot be blank.</li>"
            End If
            End Select

    End Select
    
    If txt136.Value = "" Or IsNull(txt136.Value) Then
        msgStr = msgStr & "<li><b>Batch</b> cannot be blank.</li>"
    End If
    If txt134.Value = "" Or IsNull(txt134.Value) Then
        msgStr = msgStr & "<li><b>Inventory #</b> cannot be blank.</li>"
    End If
    If txt37.Value = "" Or IsNull(txt37.Value) Then
        msgStr = msgStr & "<li><b>FFM Assign Policy # - FFM Value</b> cannot be blank.</li>"
    End If
    If txt25.Value = "" Or IsNull(txt25.Value) Then
        msgStr = msgStr & "<li><b>FFM Subscriber ID - FFM Value</b> cannot be blank.</li>"
    End If
    If txt139.Value = "" Or IsNull(txt139.Value) Then
        msgStr = msgStr & "<li><b>Benefit Year - FFM Value</b> cannot be blank.</li>"
    End If
    If txtNotes.Value = "" Or IsNull(txtNotes.Value) Then
        msgStr = msgStr & "<li><b>Notes</b> cannot be blank.</li>"
    End If


    If msgStr = "" Then 'no errors
        txtErrorBox.Value = Null
        txtErrorBar.Value = Null
        txtErrorBar.BackColor = RGB(166, 166, 166)
        CheckSubmission = False
    Else 'errors, show msg to user
        txtErrorBox.Value = headerStr & msgStr & footerStr
        txtErrorBar.Value = "Submission Errors"
        txtErrorBar.BackColor = RGB(255, 186, 0)
        CheckSubmission = True
    End If

End Function

