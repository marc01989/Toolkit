VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MC400 Import"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

 Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Form_Timer()

'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'timer set to run this sub every 2.5 mins (150000 ms)
'if now() time is greater than time report extracts drop (generally 6:49 a.m.), the import the reports and set timer back to 0 to end
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Me.TimerInterval = 10000

If TimeValue(Now()) >= #6:51:00 AM# Then
'If TimeValue(Now()) >= #10:18:00 AM# Then

    ' run job, then turn off the timer
    
    Call importExtracts
    
    Me.TimerInterval = 0
    
    Call runQueries
    
End If
    
End Sub

Public Sub importExtracts()

'timed import
'spec name = "Extract Spec"

Dim filePath As String
Dim table(1 To 4) As Variant

table(1) = "IDW Report-221910 On Exchange Active"
table(2) = "IDW Report-221910 On Exchange Termed"
table(3) = "IDW Report-235827 On Exchange Active 2019"
table(4) = "IDW Report-235827 On Exchange Termed 2019"

DoCmd.SetWarnings False

For i = 1 To 4

    filePath = "X:\Member Enrollment\Member Enrollment(DeptUsers)\Marketplace\834 Database\Extracts\" & table(i) & ".csv"

    DoCmd.runSQL ("DELETE * FROM [" & table(i) & "];")

    DoCmd.TransferText acImportDelim, "Extract Spec", table(i), filePath, 1

    'Sleep (6000)

Next i

End Sub

Public Sub runQueries()



End Sub

Private Sub btnImportExtract_Click()

'manual import
'spec name = "Extract Spec"

Dim filePath As String
Dim table(1 To 4) As Variant

table(1) = "IDW Report-221910 On Exchange Active"
table(2) = "IDW Report-221910 On Exchange Termed"
table(3) = "IDW Report-235827 On Exchange Active 2019"
table(4) = "IDW Report-235827 On Exchange Termed 2019"

DoCmd.SetWarnings False

'For i = 1 To 4

    filePath = "X:\Member Enrollment\Member Enrollment(DeptUsers)\Marketplace\834 Database\Extracts\" & table(i) & ".csv"

    DoCmd.runSQL ("DELETE * FROM [" & table(i) & "];")

    DoCmd.TransferText acImportDelim, "Extract Spec", table(i), filePath, 1

    'Sleep (6000)

'Next i


End Sub



