VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Landing Page"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub Form_Load()
    dtPickerStart.Value = Date
    dtPickerEnd.Value = Date
End Sub


Private Sub brnViewEnrollmentReports_Click()

Select Case optEnrollmentReports.Value
Case 1
    DoCmd.OpenQuery "qryTotal_LOB", acViewNormal, acReadOnly
Case 2
    DoCmd.OpenQuery "qryTotal_ChatTimeLOB", acViewNormal, acReadOnly
Case 3
    DoCmd.OpenQuery "qryTotal_LOBSpecialist", acViewNormal, acReadOnly
End Select

End Sub


Private Sub btnViewMSReports_Click()

Select Case optMSReports.Value
Case 1
DoCmd.OpenQuery "qryTotal_LOBSpecialistMS", acViewNormal, acReadOnly
End Select

End Sub

Private Sub btnTrends_Click()

Select Case optTrends.Value
Case 1
DoCmd.OpenQuery "qryUnansweredChats", acViewNormal, acReadOnly
End Select

End Sub


Private Sub btnOpenTranscripts_Click()

'open form as button
DoCmd.OpenForm "Transcript", acNormal, , , acFormPropertySettings

End Sub

Private Sub btnAdmin_Click()

DoCmd.OpenForm "Admin", acNormal, , , acFormPropertySettings

End Sub

Private Sub btnTrending_Click()

DoCmd.OpenForm "Trending", acNormal, , , acFormPropertySettings

End Sub



