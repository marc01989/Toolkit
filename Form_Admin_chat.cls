VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Admin_chat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnImport_Click()

'must have "MS Office 2016 Object Library" selected to use FileDialog object

Dim fDialog As FileDialog
Dim excelApp As Object
Dim excelWkb As Object
Dim db As DAO.Database
Dim qd As DAO.QueryDef
Dim filePath As String, fileLocation As String
Dim rowCount As Integer

Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
Set excelApp = CreateObject("excel.application")
excelApp.Visible = False
Excel.Application.ScreenUpdating = False

fileLocation = "X:\Member Enrollment\Member Enrollment(DeptUsers)\CHAT\Chat Transcripts\"

'open file dialog window
With fDialog
    .AllowMultiSelect = False
    .Title = "Select a file to import"
    .InitialFileName = fileLocation
    .Filters.Clear
        If .Show = 0 Then
            End
        Else
            filePath = .SelectedItems(1)
        End If
End With

If MsgBox("Do you want to import the selected file?", vbQuestion + vbYesNo, "Import?") = vbYes Then

    'delete old data from table
    DoCmd.runSQL ("DELETE * FROM [Chat Evaluation Report]")
    
    'run appropriate import depending on filetype selected
    Set excelWkb = excelApp.Workbooks.Open(filePath, True, True)
    
    If optFileType.Value = 1 Then
        rowCount = excelApp.Application.CountA(excelWkb.Worksheets(1).Range("A2:A20000"))
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "Chat Evaluation Report", filePath, True, "Chat Evaluation Report" & "$A2:O" & rowCount
    ElseIf optFileType.Value = 2 Then
        DoCmd.TransferText acImportDelim, "Chat Import Spec", "Chat Evaluation Report", filePath, 1
    End If
    excelWkb.Close: Set excelWkb = Nothing
    excelApp.Quit: Set excelApp = Nothing
    
    'run saved query to pull chats off [Chat Evaluation Report] table and append to [chat_data]
    DoCmd.SetWarnings True
    Set db = CurrentDb
    Set qd = db.QueryDefs("qryInsert_NewChats")
    qd.Execute
    
    qd.Close: Set qd = Nothing
    db.Close: Set db = Nothing
    
    MsgBox ("Import Complete")
    Exit Sub
Else
    End
End If

End Sub

Private Sub btnRunScript_Click()

Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim result As String
Dim rowCount As Long, rowCount2 As Long, i As Long
Dim idArray As Variant

'establish ADO connection
Set conn = CurrentProject.Connection

'open rs2 as list of distinct connid's from current chat report
Set rs2 = conn.Execute("SELECT DISTINCT Connid FROM [Chat Evaluation Report];")
rs2.Close

'open again as static pointer to get record count and save connid list into array
rs2.Open "SELECT DISTINCT Connid FROM [Chat Evaluation Report];", , adOpenStatic
rowCount = rs2.recordCount
idArray = rs2.GetRows(rowCount)

'open new rs and loop through each distinct connid, selecting specialist name and chat text for each line
DoCmd.SetWarnings False
    For i = 0 To rowCount - 1
        Set rs = New ADODB.Recordset
        rs.Open "SELECT Speakername, Chattext FROM [Chat Evaluation Report] WHERE Connid = '" & idArray(0, i) & "' ORDER BY [Msg Seq];", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rowCount2 = rs.recordCount
    
        rs.MoveFirst
        
            'iterate through each specialist name in table, adding bold html tags around name
            For j = 0 To rowCount2 - 1
                If rs.Supports(adUpdate) Then
                rs.Fields(0) = "<b>" & rs.Fields(0) & "</b>"
                rs.Update
                rs.MoveNext
                End If
            Next j
        
        'use GetString to concat rows into single string and append to transcripts table with distinct connid
        rs.MoveFirst
        result = rs.GetString(2, -1, ": ", "<br>")
        DoCmd.runSQL ("INSERT INTO chat_transcripts (Connid, transcript) VALUES ('" & idArray(0, i) & "','" & cleanText(result) & "');")
    Next i

rs2.Close: Set rs2 = Nothing
rs.Close: Set rs = Nothing
conn.Close: Set conn = Nothing

MsgBox ("Report Formatting Complete")


End Sub

Private Sub Form_Load()
dtPickerStart.Value = Date
dtPickerEnd.Value = Date
End Sub

Private Sub btnRunReport_Click()

Dim db As DAO.Database
Dim qd As DAO.QueryDef
Dim rst As DAO.Recordset

Dim fFile As Long
Dim strFile As String
Dim strString As String, dtStart As String, dtEnd As String, dtStartFile As String, dtEndFile As String
Dim qryArray(0 To 3) As String
qryArray(0) = "qryTotal_LOB"
qryArray(1) = "qryTotal_ChatTimeLOB"
qryArray(2) = "qryTotal_LOBSpecialist"
qryArray(3) = "qryTotal_LOBSpecialistMS"

dtStart = Format(dtPickerStart.Value, "mm/dd/yyyy")
dtEnd = Format(dtPickerEnd.Value, "mm/dd/yyyy")

Dim qryName(0 To 3) As String
qryName(0) = "Chat Data from " & dtStart & " to " & dtEnd
qryName(1) = ""
qryName(2) = ""
qryName(3) = ""

fFile = FreeFile

'Set output text file path and name
strFile = "X:\Member Enrollment\Member Enrollment(DeptUsers)\CHAT\Database\Reports\Chat_Reports_" & Replace(dtStart, "/", ".") & " - " & Replace(dtEnd, "/", ".") & ".txt"
Open strFile For Output As #fFile

Set db = CurrentDb

For j = 0 To 3

    Set qd = db.QueryDefs(qryArray(j))
    
    'refresh parameters of each query
    qd.Parameters(0).Value = dtStart
    qd.Parameters(1).Value = dtEnd
    
    Set rst = qd.OpenRecordset
    
'           test = qd.sql
'            Print #fFile, test

    If rst Is Nothing Then Exit Sub
        
        If Not rst.EOF And Not rst.BOF Then

        Debug.Print rst.recordCount
        
        'print name of query and date range before the data
        Print #fFile, qryName(j)
        
        'Store header
             For i = 0 To rst.Fields.Count - 1
                  strString = strString & rst.Fields(i).Name & vbTab
                     'Debug.Print rst.Fields(i).Name
             Next i
             Print #fFile, strString
             'Reset the string to empty for the next record
              strString = ""
              i = 0
        
        'Write Records to file
        
           Do Until rst.EOF
               'Loop across all fields in recordset delimiting them with a tab key
                For i = 0 To rst.Fields.Count - 1
                    strString = strString & rst(i) & vbTab
                Next i
                'Output the resulting string to the text file
                Print #fFile, strString
                'Reset the string to empty for the next record
                strString = ""
                'Move to the next record in the recordset
                rst.MoveNext
           Loop
           
           'Close the recordset
           rst.Close: Set rst = Nothing
           qd.Close: Set qd = Nothing
            
        End If
    
Next j
    
Close #fFile

db.Close: Set db = Nothing

If MsgBox("Report created. Do you want to view the file?", vbYesNo, "Success") = vbYes Then
    Shell "explorer.exe " & "X:\Member Enrollment\Member Enrollment(DeptUsers)\CHAT\Database\Reports\", vbNormalFocus
End If


End Sub


