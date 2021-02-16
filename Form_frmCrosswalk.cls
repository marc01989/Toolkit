VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCrosswalk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim criteria As String
Dim is_crosswalked As Integer


'-----------------------------------------------------------------------------------------------------------------------------------------------------------
'crosswalks plan data from 2018 --> 2019
'plans are matched on qhpid which is housed in junction_18_19 table
'-----------------------------------------------------------------------------------------------------------------------------------------------------------


Private Sub Form_Load()
    txtIsCrawford.BackColor = RGB(250, 250, 250)
End Sub

Private Sub btnSearchCrosswalk_Click()

    If Not IsNull(txtQHPID.Value) Then
        Call searchQHPID
    ElseIf Not IsNull(txtGroup.Value) And Not IsNull(txtSubGroup.Value) Then
        Call searchGroupSubGroup
    ElseIf (IsNull(txtGroup.Value) And Not IsNull(txtSubGroup.Value)) Or (Not IsNull(txtGroup.Value) And IsNull(txtSubGroup.Value)) Then
        MsgBox ("Group and Subgroup boxes must contain values before searching")
        Exit Sub
    ElseIf IsNull(txtGroup.Value) And IsNull(txtSubGroup.Value) And IsNull(txtQHPID.Value) Then
        MsgBox ("Please enter a search value")
        Exit Sub
    End If

End Sub

Public Sub searchQHPID()

    Call clearBoxes
    
    
    'check if QHPID is valid format
    If checkQHPID(txtQHPID.Value) = False Then
        MsgBox ("Invalid QHPID - please check and try again")
        Exit Sub
    End If
    
    'find QHPID from lookup tables and return to insert into criteria string
    'run query subroutine with criteria
    
    If optYear.Value = 2 Then '2018
        If getQHPID(optYear.Value, txtQHPID.Value) = True Then
            criteria = "([junction_18_19].[2018_qhpid] = '" & txtQHPID.Value & "')"
            Call runSQL(criteria)
        Else
            If MsgBox("Plan is not crosswalked." & vbCrLf & vbCrLf & "Please use the single year search from the Helpful Links section.", vbInformation + vbOKOnly) = vbOK Then
                Call clearBoxes
                Set rs = Nothing
                Exit Sub
            End If
        End If
    ElseIf optYear.Value = 3 Then '2019
        If getQHPID(optYear.Value, txtQHPID.Value) = True Then
            criteria = "([junction_18_19].[2019_qhpid] = '" & txtQHPID.Value & "')"
            Call runSQL(criteria)
        Else
            If MsgBox("Plan is not crosswalked." & vbCrLf & vbCrLf & "Please use the single year search from the Helpful Links section.", vbInformation + vbOKOnly) = vbOK Then
                Call clearBoxes
                Set rs = Nothing
                Exit Sub
            End If
        End If
    End If

End Sub

Public Sub searchGroupSubGroup()

    Call clearBoxes
    
    If optYear.Value = 3 Then '2019
        If getGroupSubgroup(optYear.Value, txtGroup.Value, txtSubGroup.Value) = True Then
            criteria = "([2019_group] = '" & txtGroup.Value & "' AND [2019_subgroup] = '" & txtSubGroup.Value & "')"
            Call runSQL(criteria)
        Else
            If MsgBox("Plan is not crosswalked." & vbCrLf & vbCrLf & "Please use the single year search from the Helpful Links section.", vbInformation + vbOKOnly) = vbOK Then
                Call clearBoxes
                Set rs = Nothing
                Exit Sub
            End If
        End If
    ElseIf optYear.Value = 2 Then '2018
        If getGroupSubgroup(optYear.Value, txtGroup.Value, txtSubGroup.Value) = True Then
            criteria = "([2018_group] = '" & txtGroup.Value & "' AND [2018_subgroup] = '" & txtSubGroup.Value & "')"
            Call runSQL(criteria)
        Else
            If MsgBox("Plan is not crosswalked." & vbCrLf & vbCrLf & "Please use the single year search from the Helpful Links section.", vbInformation + vbOKOnly) = vbOK Then
                Call clearBoxes
                Set rs = Nothing
                Exit Sub
            End If
        End If
    End If


End Sub

Public Sub runSQL(criteria As String)

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT DISTINCT * " & _
        " FROM 2018_qhpid_tbl INNER JOIN (2019_qhpid_tbl INNER JOIN junction_18_19 ON [2019_qhpid_tbl].[2019_qhpid]=junction_18_19.[2019_qhpid]) ON ([2018_qhpid_tbl].[2018_qhpid]=junction_18_19.[2018_qhpid]) AND ([2018_qhpid_tbl].[2018_dental_buy_up]=[2019_qhpid_tbl].[2019_dental_buy_up] Or [2018_qhpid_tbl].[2018_dental_buy_up] Is Null And [2019_qhpid_tbl].[2019_dental_buy_up] Is Null) " & _
        " WHERE " & criteria & _
        " ORDER BY [2019_qhpid_tbl].[2019_qhpid], is_crawford, [2018_subgroup];", dbOpenSnapshot, dbReadOnly)
    
    If rs.recordCount > 0 Then
        rs.MoveFirst
        rs.MoveLast
        rs.MoveFirst
        
        is_crosswalked = 1
        MsgBox (rs.recordCount & " crosswalked plan(s) have been found")
        Call populateBoxes
    
    Else
        If MsgBox("Plan is not crosswalked - please try the single-year search.", vbInformation + vbOKOnly) = vbOK Then
            Call clearBoxes
            Exit Sub
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
        
        is_crosswalked = 0
        MsgBox ("Plan is not crosswalked. " & vbCrLf & vbCrLf & rs.recordCount & " plan(s) have been found")
        
        If year = "2018" Then
            Call pop2018
        ElseIf year = "2019" Then
            Call pop2019
        End If
    
    Else
        If MsgBox("Plan not found - please refine your search and try again.", vbInformation + vbOKOnly) = vbOK Then
            Call clearBoxes
            Exit Sub
        End If
    End If

End Sub

Private Sub btnLeft_Click()

    If criteria = "" Or IsNull(criteria) Then
    Else
        rs.FindPrevious (criteria)
        If is_crosswalked = 1 Then
            Call populateBoxes
        ElseIf s_crosswalked = 0 Then
            Call pop2019
        End If
    End If

End Sub

Private Sub btnRight_Click()

    If criteria = "" Or IsNull(criteria) Then
    Else
        rs.FindNext (criteria)
        If is_crosswalked = 1 Then
            Call populateBoxes
        ElseIf s_crosswalked = 0 Then
            Call pop2019
        End If
    End If

End Sub

Public Sub populateBoxes()

    txt2018QHPID.Value = rs![junction_18_19.2018_qhpid]
    txt2018Group.Value = rs![2018_group]
    txt2018Subgroup.Value = rs![2018_subgroup]
    txt2018PlanName.Value = rs![2018_plan_name]
    txt2018PlanCode.Value = rs![2018_plan_code]
    txt2018EHBRider.Value = rs![2018_ehb_rider]
    txt2018PlanType.Value = rs![2018_plan_type]
    txt2018RxRider.Value = rs![2018_rx_rider]
    txt2018VisionRider.Value = rs![2018_vision_rider]
    txt2018DentalRider.Value = rs![2018_dental_buy_up]
    
    txt2019QHPID.Value = rs![junction_18_19.2019_qhpid]
    txt2019Group.Value = rs![2019_group]
    txt2019Subgroup.Value = rs![2019_subgroup]
    txt2019PlanName.Value = rs![2019_plan_name]
    txt2019PlanCode.Value = rs![2019_plan_code]
    txt2019EHBRider.Value = rs![2019_ehb_rider]
    txt2019PlanType.Value = rs![2019_plan_type]
    txt2019RxRider.Value = rs![2019_rx_rider]
    txt2019VisionRider.Value = rs![2019_vision_rider]
    txt2019DentalRider.Value = rs![2019_dental_buy_up]
    
    
    If rs![is_crawford] = 1 Then
        txtIsCrawford.Value = "Crawford County!"
        txtIsCrawford.BackColor = RGB(238, 238, 0)
    Else
        txtIsCrawford.Value = Null
        txtIsCrawford.BackColor = RGB(250, 250, 250)
    End If
    
    txtRecordCount.Value = rs.recordCount
    Call QHPIDMatch

End Sub

Public Sub pop2018()

    txt2018QHPID.Value = rs![2018_qhpid]
    txt2018Group.Value = rs![2018_group]
    txt2018Subgroup.Value = rs![2018_subgroup]
    txt2018PlanName.Value = rs![2018_plan_name]
    txt2018PlanCode.Value = rs![2018_plan_code]
    txt2018EHBRider.Value = rs![2018_ehb_rider]
    txt2018PlanType.Value = rs![2018_plan_type]
    txt2018RxRider.Value = rs![2018_rx_rider]
    txt2018VisionRider.Value = rs![2018_vision_rider]
    txt2018DentalRider.Value = rs![2018_dental_buy_up]
    
    txtRecordCount.Value = rs.recordCount

End Sub

Public Sub pop2019()

    txt2019QHPID.Value = rs![2019_qhpid]
    txt2019Group.Value = rs![2019_group]
    txt2019Subgroup.Value = rs![2019_subgroup]
    txt2019PlanName.Value = rs![2019_plan_name]
    txt2019PlanCode.Value = rs![2019_plan_code]
    txt2019EHBRider.Value = rs![2019_ehb_rider]
    txt2019PlanType.Value = rs![2019_plan_type]
    txt2019RxRider.Value = rs![2019_rx_rider]
    txt2019VisionRider.Value = rs![2019_vision_rider]
    txt2019DentalRider.Value = rs![2019_dental_buy_up]
    
    txtRecordCount.Value = rs.recordCount

End Sub

Public Sub clearBoxes()

Dim ctrl As Control

    txt2019PlanName.Value = Null
    txt2019PlanCode.Value = Null
    txt2019EHBRider.Value = Null
    txt2019DentalRider.Value = Null
    txt2019PlanType.Value = Null
    txt2019RxRider.Value = Null
    txt2019VisionRider.Value = Null
    
    txt2019QHPID.Value = Null
    txt2019Group.Value = Null
    txt2019Subgroup.Value = Null
    
    txt2018QHPID.Value = Null
    txt2018Group.Value = Null
    txt2018Subgroup.Value = Null
    
    txt2018PlanName.Value = Null
    txt2018PlanCode.Value = Null
    txt2018EHBRider.Value = Null
    txt2018DentalRider.Value = Null
    txt2018PlanType.Value = Null
    txt2018RxRider.Value = Null
    txt2018VisionRider.Value = Null
    
    txtIsCrawford.Value = Null
    txtIsCrawford.BackColor = RGB(250, 250, 250)
    txtRecordCount.Value = Null
    txtQHPIDMatch.Value = Null

End Sub

Public Sub QHPIDMatch()

    If txt2018QHPID.Value = txt2019QHPID.Value Then
        txtQHPIDMatch.Value = "Same QHPID"
    Else
        txtQHPIDMatch.Value = "Different QHPID"
    End If

End Sub


'control actions
Private Sub txtQHPID_KeyDown(KeyCode As Integer, Shift As Integer)
    txtGroup.Value = Null
    txtSubGroup.Value = Null
End Sub

Private Sub txtSubGroup_KeyDown(KeyCode As Integer, Shift As Integer)
    txtQHPID.Value = Null
End Sub
Private Sub txtGroup_KeyDown(KeyCode As Integer, Shift As Integer)
    txtQHPID.Value = Null
End Sub

Private Sub btnSingleYear_click()
    DoCmd.OpenForm "frmSingleYearSearch", acNormal, , , acFormPropertySettings
End Sub

Private Sub btnInfo_Click()
    Application.FollowHyperlink "https://www.upmchealthplan.com/members/networkmap/"
End Sub

