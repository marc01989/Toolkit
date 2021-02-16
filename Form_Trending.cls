VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Trending"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cboSearchCategories_AfterUpdate()

Select Case cboSearchCategories.Value
    
    Case 1 'pharmacy
        txtSearch.Value = "'Pharmacy', 'ESI', 'ESD', 'Prescription', 'Express Scripts'"
    Case 2 'reinstates
        txtSearch.Value = "'Reinstate', 'Reinstated', 'Ops', 'Operations'"
    Case 3 'Terminations
        txtSearch.Value = "'Term', 'Termination', 'Terminate'"
    Case 4 'HICs
        txtSearch.Value = "'HICs', 'HIX', 'Appeal', 'CMS'"
    Case 5 'cancelations
        txtSearch.Value = "'Cancel', 'Cancelled', 'Void', 'Voided'"
    
End Select

End Sub

Private Sub Form_Load()
    dtPickerStart.Value = Date
    dtPickerEnd.Value = Date
End Sub

Private Sub btnSearch_Click()

Dim searchStr As String

Select Case cboSearchCategories.Value
    Case 1 'pharmacy
        searchStr = "\b(?:esi|pharmacy|esd|prescription|Express Scripts)\b"
    Case 2 'reinstates
        searchStr = "\b(?:reinstate|reinstated|ops|operations)\b"
    Case 3 'Terminations
        searchStr = "\b(?:term|termination|terminate)\b"
    Case 4 'HICs
        searchStr = "\b(?:hics|hix|appeal|cms)\b"
    Case 5 'cancelations
        searchStr = "\b(?:cancel|cancelled|void|voided)\b"
End Select

'Select Case optSearchType
'    Case 1 '--in text--
'        searchStr = txtSearch.Value
'    Case 2 '--not in text--
'        searchStr = "^((?!" & txtSearch.Value & ").)*$"
'End Select

Call RegularExpression(searchStr)

End Sub

Public Sub RegularExpression(searchStr As String)

Dim regex As RegExp
Dim colRegMatch As MatchCollection
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim resultStr As String, dateStart As String, dateEnd As String
Dim counter As Integer

dateStart = Format(dtPickerStart.Value, "yyyymmdd")
dateEnd = Format(dtPickerEnd.Value, "yyyymmdd")

'instantiate regex object, set properties
'set pattern - what you are/aren't looking for
'find string where this expression is not present:  ^((?! PHRASE HERE  ).)*$
Set regex = New RegExp
With regex
    .MultiLine = False
    .Global = True
    .IgnoreCase = True
    .Pattern = searchStr
End With

resultStr = ""
counter = 0

'open recordset to iterate through
Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT chat_transcripts.transcript, [Chat Start Time], [Chat End Time], [Chat Time in Mins] " & _
" FROM chat_data INNER JOIN chat_transcripts ON chat_data.Connid = chat_transcripts.Connid " & _
" WHERE FORMAT(LEFT([Chat Start Time], 10), 'yyyymmdd') >= '" & dateStart & "' AND FORMAT(LEFT([Chat Start Time], 10), 'yyyymmdd') <= '" & dateEnd & "'" & _
" ORDER BY FORMAT(LEFT([Chat Start Time], 10), 'yyyymmdd') ;")

'with recordset, test each record against the pattern in regex
With rs
    If .recordCount > 0 Then
        .MoveFirst
        Do While Not .EOF
            Dim tempStr As String
'if pattern is found, save the record to string with HTML formatting
                If regex.test(.Fields(0)) = True Then
                    tempStr = .Fields(0)
                    tempStr = regex.Replace(tempStr, "[<i><b>$&</b></i>]")
                    'Debug.Print tempStr
                    
                    counter = counter + 1
                    'Set colRegMatch = regex.Execute(.Fields(0))
                    'Debug.Print .Fields(0) & "<br><br>"
                    resultStr = resultStr & "<b>" & counter & ".) </b><br>" & _
                    "<b>Date: </b>" & Left(.Fields(1), 10) & "<br>" & _
                    "<b>Time: </b>" & Right(.Fields(1), 8) & " to " & Right(.Fields(2), 8) & "<br>" & _
                    "<B>Total time: </b>" & .Fields(3) & " minutes <br><br>" & _
                    tempStr & "<br>"
                End If
            .MoveNext
        Loop
    End If
End With

rs.Close: Set rs = Nothing
db.Close: Set db = Nothing

MsgBox "no errors"

txtResults.Value = Null
txtResults.Value = resultStr
txtChatCount.Value = Null
txtChatCount.Value = counter

End Sub
