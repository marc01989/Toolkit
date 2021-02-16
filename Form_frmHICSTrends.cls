VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmHICSTrends"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cboDate1_AfterUpdate()
    cboDate2.Requery
End Sub
'search btn, returns case narrative
Private Sub Command11_Click()
    Dim db As Database
    Dim rs As Recordset
    Set db = CurrentDb
    Set rs = db.OpenRecordset("Select * From tblCMSCaseData Where case_id = " & "'" & txtCaseID.Value & "';", dbOpenSnapshot)
        If rs.EOF = False Then
        txtNarrative.Value = rs![case_narrative]
        rs.Close: Set rs = Nothing: db.Close: Set db = Nothing
        Else: rs.Close: Set rs = Nothing: db.Close: Set db = Nothing: MsgBox "no case found"
        End If
        
End Sub
'buttons
Private Sub Command24_Click()
    DoCmd.OpenQuery "qryTrendsByCaseID", acViewNormal, acReadOnly
    
End Sub

Private Sub Command25_Click()
    DoCmd.OpenQuery "qryTransferredCasesByCategory", acViewNormal, acReadOnly
End Sub

Private Sub Command6_Click()
    DoCmd.OpenQuery "qrySubcategoryTrends", acViewNormal, acReadOnly
End Sub
