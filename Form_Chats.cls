VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Chats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim arrayChats As Variant
Dim recordMax As Long, recordCounter As Long
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim strCon As String


'--REFERENCES NEEDED FOR ADO CONNECTION
    '-Microsoft Activex Data Objects 6.0 Library
    
Private Sub Form_Load()
    '--START CONNECTION HERE--
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    strCon = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
    "Data Source=X:\Member Enrollment\Member Enrollment(DeptUsers)\CHAT\Database\Chat Database.accde;" & _
    "User Id=admin;Password="
    conn.Open (strCon)
    
End Sub
Private Sub Form_Close()
    '--KILL CONNECTION--
    Set rs = Nothing
    conn.Close: Set conn = Nothing
End Sub
    
Private Sub btnSearch_Click()

    Dim sql As String
    Dim weekStart As String, weekEnd As String, chatTable As String
    
    'check that required controls aren't null
    If week_id.Value = "" Or IsNull(week_id.Value) Then
        MsgBox "Must select date before searching"
    End If
    If IsNull(cboSpecialist.Value) Or cboSpecialist.Value = "" Then
        Exit Sub
    End If
    
    'set date range
    weekStart = Format(week_id.Column(1), "yyyy-mm-dd")
    weekEnd = Format(DateAdd("d", 6, weekStart), "yyyymmdd")
    weekStart = Format(weekStart, "yyyymmdd")
    
    sql = "SELECT chat_transcripts.Connid, transcript, [Chat Start Time], [Chat End Time], Caller, [Agent ID], MemberID " & _
    " FROM chat_data LEFT JOIN chat_transcripts ON chat_data.Connid = chat_transcripts.Connid " & _
    " WHERE ([Agent ID] = '" & cboSpecialist.Value & "' AND FORMAT((Left([Chat Start Time],10)), 'YYYYMMDD') >= '" & weekStart & "' AND FORMAT((Left([Chat Start Time],10)), 'YYYYMMDD') < '" & weekEnd & "')" & _
    " ORDER BY [Chat Start Time] ASC;"
    
    
On Error GoTo err1:

    '--OPEN CONNECTION HERE--
    rs.Open sql, conn, adOpenKeyset
    
    With rs
    'if records are returned, save into array, set recordMax and recordCounter
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
    'if no records, set everything to default values. recordMax set to 0 means no records, disables controls
            MsgBox "No records found"
            Call clearBoxes
            arrayChats = Empty
            recordMax = 0
            recordCounter = 0
        End If
    End With
    rs.Close

    
err1:
    Select Case Err.Number
        Case 0
        Case Else
            Call LogError(Err.Number & " " & Err.Description, "Chats; btnSearch_Click()")
            If MsgBox("Error connecting to database. See admin for assistance.", vbCritical + vbOKOnly, "System Error") = vbOK Then: Exit Sub
            Exit Sub
    End Select
    
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
Private Sub Label100_Click()
    DoCmd.OpenForm "Home", acNormal, , , acFormPropertySettings
End Sub

Private Sub Label101_Click()
    DoCmd.Close acForm, "Chats", acSaveNo
End Sub
