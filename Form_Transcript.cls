VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Transcript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim arrayChats As Variant
Dim recordMax As Long, recordCounter As Long

Private Sub Form_Load()

cboWeek.Value = Null
cboWeek.RowSource = ""
cboSpecialist.Value = Null
cboSpecialist.RowSource = ""
Call optYear_Click

End Sub

Private Sub cboDepartment_AfterUpdate()

cboSpecialist.RowSource = "SELECT Name, EmployeeID FROM department INNER JOIN specialists ON department.department_id = specialists.department_id WHERE specialists.department_id = " & cboDepartment.Value & " ORDER BY Name;"
cboSpecialist.Value = Null

End Sub

Private Sub btnSearch_Click()

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim criteria As String

If CheckSubmission = True Then
    If MsgBox("Boxes cannot be empty. See Submission Errors for more info.", vbExclamation + vbOKOnly, "Cannot Submit") = vbOK Then: Exit Sub
End If

weekStart = Format(cboWeek.Column(1), "yyyy-mm-dd")
weekEnd = Format(DateAdd("d", 6, weekStart), "yyyymmdd")
weekStart = Format(weekStart, "yyyymmdd")
departmentId = cboDepartment.Value
nameId = cboSpecialist.Value


If nameId = "" Or IsNull(nameId) Then

    '--SEARCH BY DEPARTMENT (name is null)
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT chat_transcripts.Connid, chat_transcripts.transcript, chat_data.[Chat Start Time], chat_data.[Chat End Time], chat_data.Caller, chat_data.[Agent ID], chat_data.MemberID " & _
    " FROM (chat_transcripts INNER JOIN chat_data ON chat_transcripts.Connid = chat_data.Connid) INNER JOIN (department INNER JOIN specialists ON department.department_id = specialists.department_id) ON chat_data.[Agent ID] = specialists.EmployeeID " & _
    " WHERE (specialists.department_id = " & departmentId & " AND FORMAT((Left([Chat Start Time],10)), 'YYYYMMDD') >= '" & weekStart & "' AND FORMAT((Left([Chat Start Time],10)), 'YYYYMMDD') < '" & weekEnd & "') " & _
    " ORDER BY [Chat Start Time] ASC")
Else
    '--SEARCH BY NAME
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT chat_transcripts.Connid, transcript, [Chat Start Time], [Chat End Time], Caller, [Agent ID], MemberID " & _
        " FROM chat_data LEFT JOIN chat_transcripts ON chat_data.Connid = chat_transcripts.Connid " & _
        " WHERE ([Agent ID] = '" & cboSpecialist.Column(1) & "' AND FORMAT((Left([Chat Start Time],10)), 'YYYYMMDD') >= '" & weekStart & "' AND FORMAT((Left([Chat Start Time],10)), 'YYYYMMDD') < '" & weekEnd & "')" & _
        " ORDER BY [Chat Start Time] ASC;")
End If

criteria = "([Agent ID] = '" & cboSpecialist.Column(1) & "' AND FORMAT((Left([Chat Start Time],10)), 'YYYYMMDD') >= '" & weekStart & "' AND FORMAT((Left([Chat Start Time],10)), 'YYYYMMDD') < '" & weekEnd & "')"
    
    
With rs
    If .recordCount > 0 Then
        .MoveFirst
        .MoveLast
        .MoveFirst
        txtChatCount.Value = .recordCount
        recordMax = .recordCount - 1
        recordCounter = 0
        arrayChats = .GetRows(.recordCount)
        Call populateBoxes(recordCounter)
    Else
        MsgBox "No records found"
        Call clearBoxes
        arrayChats = Empty
        recordMax = 0
        recordCounter = 0
    End If
End With

rs.Close: Set rs = Nothing
db.Close: Set db = Nothing

End Sub

Public Sub populateBoxes(i As Long)

    txtConnid.Value = arrayChats(0, i) 'connid
    txtChatStartTime.Value = arrayChats(2, i) 'chatstart
    txtChatEndTime.Value = arrayChats(3, i) 'chatend
    txtCaller.Value = arrayChats(4, i) 'caller - MS rep
    txtAgentID.Value = arrayChats(5, i) 'receiver - Enrollment
    txtMemberID.Value = arrayChats(6, i) 'member id
    txtTranscript.Value = arrayChats(1, i) 'transctipt

End Sub

Public Sub clearBoxes()

    
    txtConnid.Value = Null
    txtChatStartTime.Value = Null
    txtChatEndTime.Value = Null
    txtCaller.Value = Null
    txtAgentID.Value = Null
    txtMemberID.Value = Null
    txtTranscript.Value = Null
    txtChatCount.Value = Null


End Sub

Private Sub btnForward_Click()

If recordMax = 0 Then Exit Sub

If recordCounter = recordMax Then
    Exit Sub
Else
    recordCounter = recordCounter + 1
    Call populateBoxes(recordCounter)
End If
End Sub

Private Sub btnBack_Click()
If recordMax = 0 Then Exit Sub

If recordCounter = 0 Then
    Exit Sub
Else
    recordCounter = recordCounter - 1
    Call populateBoxes(recordCounter)
End If

End Sub

Private Sub optYear_Click()

'changes contents of week dropdown box when selected

If optYear.Value = 1 Then
    cboWeek.RowSource = "SELECT [week_tbl].[week_id], [week_tbl].[week_start] FROM week_tbl WHERE week_id >= 54 AND week_id < 106;"
    cboWeek.Value = Null
ElseIf optYear.Value = 2 Then
    cboWeek.RowSource = "SELECT [week_tbl].[week_id], [week_tbl].[week_start] FROM week_tbl WHERE week_id >= 106;"
    cboWeek.Value = Null
End If
cboWeek.Value = cboWeek.Column(0, 0)

End Sub

Public Function CheckSubmission() As Boolean
 
    Dim msgStr As String, headerStr As String, footerStr As String
    headerStr = "-Submission Errors- <br><br> <ul>"
    footerStr = "</ul>"
    
    If cboDepartment.Value = "" Or IsNull(cboDepartment.Value) = True Then
        msgStr = msgStr & "<li><b>Department</b> cannot be blank.</li>"
    End If
    If cboWeek.Value = "" Or IsNull(cboWeek.Value) = True Then
        msgStr = msgStr & "<li><b>Week Of</b> cannot be blank.</li>"
    End If
    
    If msgStr = "" Then
        txtErrorBox.Value = Null
        CheckSubmission = False
    Else
        txtErrorBox.Value = headerStr & msgStr & footerStr
        CheckSubmission = True
    End If
     
End Function
