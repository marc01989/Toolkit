VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSingleYearSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim criteria As String
Dim planYear As String


Private Sub btnSingleYearSearch_Click()


If Not IsNull(txtQHPIDSearch.Value) Then

    Call searchQHPID

ElseIf Not IsNull(txtGroupSearch.Value) And Not IsNull(txtSubGroupSearch.Value) Then

    Call searchGroupSubGroup
    
ElseIf (IsNull(txtGroupSearch.Value) And Not IsNull(txtSubGroupSearch.Value)) Or (Not IsNull(txtGroupSearch.Value) And IsNull(txtSubGroupSearch.Value)) Then

    MsgBox ("Group and Subgroup boxes must contain values before searching")
    Exit Sub

ElseIf IsNull(txtGroupSearch.Value) And IsNull(txtSubGroupSearch.Value) And IsNull(txtQHPIDSearch.Value) Then

    MsgBox ("Please enter a search value")
    Exit Sub

End If


End Sub

Public Sub searchQHPID()

Call clearBoxes


'check if QHPID is valid format
If checkQHPID(txtQHPIDSearch.Value) = False Then
    MsgBox ("Invalid QHPID - please check and try again")
    Exit Sub
End If

'find QHPID from lookup tables and return to insert into criteria string
'run query subroutine with criteria

If optYear.Value = 1 Then '2017
    If getQHPIDSingle(optYear.Value, txtQHPIDSearch.Value) = True Then
        criteria = "([2017_qhpid] = '" & txtQHPIDSearch.Value & "')"
        Call runSQLSingle(criteria, "2017")
    Else
        If MsgBox("Plan not found - please refine your search and try again.", vbInformation + vbOKOnly) = vbOK Then
            Call clearBoxes
            Set rs = Nothing
            Exit Sub
        End If
    End If

ElseIf optYear.Value = 2 Then '2018
    If getQHPIDSingle(optYear.Value, txtQHPIDSearch.Value) = True Then
        criteria = "([2018_qhpid] = '" & txtQHPIDSearch.Value & "')"
        Call runSQLSingle(criteria, "2018")
    Else
        If MsgBox("Plan not found - please refine your search and try again.", vbInformation + vbOKOnly) = vbOK Then
            Call clearBoxes
            Set rs = Nothing
            Exit Sub
        End If
    End If
ElseIf optYear.Value = 3 Then '2019
    If getQHPIDSingle(optYear.Value, txtQHPIDSearch.Value) = True Then
        criteria = "([2019_qhpid] = '" & txtQHPIDSearch.Value & "')"
        Call runSQLSingle(criteria, "2019")
    Else
        If MsgBox("Plan not found - please refine your search and try again.", vbInformation + vbOKOnly) = vbOK Then
            Call clearBoxes
            Set rs = Nothing
            Exit Sub
        End If
    End If
End If

End Sub

Public Sub runSQLSingle(criteria As String, year As String)

Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT DISTINCT * " & _
    " FROM [" & year & "_qhpid_tbl] " & _
    " WHERE " & criteria & _
    " ORDER BY [" & year & "_subgroup];", dbOpenSnapshot, dbReadOnly)

If rs.recordCount > 0 Then
    rs.MoveFirst
    rs.MoveLast
    rs.MoveFirst
    
    planYear = year
    MsgBox (rs.recordCount & " plan(s) have been found")
    
    If year = "2018" Then
        Call pop2018
    ElseIf year = "2019" Then
        Call pop2019
    ElseIf year = "2017" Then
        Call pop2017
    End If

Else
    If MsgBox("Plan not found - please refine your search and try again.", vbInformation + vbOKOnly) = vbOK Then
        Call clearBoxes
        Exit Sub
    End If
End If

End Sub

Public Sub searchGroupSubGroup()

Call clearBoxes

If optYear.Value = 3 Then '2019

    If getGroupSubgroupSingle(optYear.Value, txtGroupSearch.Value, txtSubGroupSearch.Value) = True Then
        criteria = "([2019_group] = '" & txtGroupSearch.Value & "' AND [2019_subgroup] = '" & txtSubGroupSearch.Value & "')"
        Call runSQLSingle(criteria, "2019")
    Else
        If MsgBox("Plan not found - please refine your search and try again.", vbInformation + vbOKOnly) = vbOK Then
            Call clearBoxes
            Set rs = Nothing
            Exit Sub
        End If
    End If
ElseIf optYear.Value = 2 Then '2018

    If getGroupSubgroupSingle(optYear.Value, txtGroupSearch.Value, txtSubGroupSearch.Value) = True Then
        criteria = "([2018_group] = '" & txtGroupSearch.Value & "' AND [2018_subgroup] = '" & txtSubGroupSearch.Value & "')"
        Call runSQLSingle(criteria, "2018")
    Else
        If MsgBox("Plan not found - please refine your search and try again.", vbInformation + vbOKOnly) = vbOK Then
            Call clearBoxes
            Set rs = Nothing
            Exit Sub
        End If
    End If
    
ElseIf optYear.Value = 1 Then '2017

    If getGroupSubgroupSingle(optYear.Value, txtGroupSearch.Value, txtSubGroupSearch.Value) = True Then
        criteria = "([2017_group] = '" & txtGroupSearch.Value & "' AND [2017_subgroup] = '" & txtSubGroupSearch.Value & "')"
        Call runSQLSingle(criteria, "2017")
    Else
        If MsgBox("Plan not found - please refine your search and try again.", vbInformation + vbOKOnly) = vbOK Then
            Call clearBoxes
            Set rs = Nothing
            Exit Sub
        End If
    End If
End If


End Sub

Public Sub pop2017()

txtQHPID.Value = rs![2017_qhpid]
txtGroup.Value = rs![2017_group]
txtSubGroup.Value = rs![2017_subgroup]
txtMetalLevel.Value = rs![2017_metal]

txtPlanName.Value = rs![2017_plan_name]
txtPlanCode.Value = rs![2017_plan_code]
txtEHBRider.Value = rs![2017_ehb_rider]
txtDentalRider.Value = rs![2017_dental_buy_up]
txtPlanType.Value = rs![2017_plan_type]
txtRxRider.Value = rs![2017_rx_rider]
txtVisionRider.Value = rs![2017_vision_rider]

txtPlanYear.Value = rs![plan_year]
txtRecordCount.Value = rs.recordCount

End Sub

Public Sub pop2018()


txtQHPID.Value = rs![2018_qhpid]
txtGroup.Value = rs![2018_group]
txtSubGroup.Value = rs![2018_subgroup]
txtMetalLevel.Value = rs![2018_metal]

txtPlanName.Value = rs![2018_plan_name]
txtPlanCode.Value = rs![2018_plan_code]
txtEHBRider.Value = rs![2018_ehb_rider]
txtDentalRider.Value = rs![2018_dental_buy_up]
txtPlanType.Value = rs![2018_plan_type]
txtRxRider.Value = rs![2018_rx_rider]
txtVisionRider.Value = rs![2018_vision_rider]

txtPlanYear.Value = rs![plan_year]
txtRecordCount.Value = rs.recordCount

End Sub

Public Sub pop2019()

txtQHPID.Value = rs![2019_qhpid]
txtGroup.Value = rs![2019_group]
txtSubGroup.Value = rs![2019_subgroup]
txtMetalLevel.Value = rs![2019_metal]

txtPlanName.Value = rs![2019_plan_name]
txtPlanCode.Value = rs![2019_plan_code]
txtEHBRider.Value = rs![2019_ehb_rider]
txtDentalRider.Value = rs![2019_dental_buy_up]
txtPlanType.Value = rs![2019_plan_type]
txtRxRider.Value = rs![2019_rx_rider]
txtVisionRider.Value = rs![2019_vision_rider]

txtPlanYear.Value = rs![plan_year]
txtRecordCount.Value = rs.recordCount

End Sub

Public Sub clearBoxes()

txtQHPID.Value = Null
txtGroup.Value = Null
txtSubGroup.Value = Null
txtMetalLevel.Value = Null

txtPlanName.Value = Null
txtPlanCode.Value = Null
txtEHBRider.Value = Null
txtDentalRider.Value = Null
txtPlanType.Value = Null
txtRxRider.Value = Null
txtVisionRider.Value = Null

txtPlanYear.Value = Null
txtRecordCount.Value = Null

End Sub

Private Sub btnLeft_Click()

If criteria = "" Or IsNull(criteria) Then
Else
    rs.FindPrevious (criteria)
    If planYear = "2018" Then
        Call pop2018
    ElseIf planYear = "2019" Then
        Call pop2019
    ElseIf planYear = "2017" Then
        Call pop2017
    End If
End If

End Sub

Private Sub btnRight_Click()

If criteria = "" Or IsNull(criteria) Then
Else
    rs.FindNext (criteria)
    If planYear = "2018" Then
        Call pop2018
    ElseIf planYear = "2019" Then
        Call pop2019
    ElseIf planYear = "2017" Then
        Call pop2017
    End If
End If

End Sub


Private Sub txtQHPIDSearch_KeyDown(KeyCode As Integer, Shift As Integer)

txtGroupSearch.Value = Null
txtSubGroupSearch.Value = Null

End Sub

Private Sub txtSubGroupSearch_KeyDown(KeyCode As Integer, Shift As Integer)

txtQHPIDSearch.Value = Null

End Sub


Private Sub txtGroupSearch_KeyDown(KeyCode As Integer, Shift As Integer)

txtQHPIDSearch.Value = Null

End Sub

