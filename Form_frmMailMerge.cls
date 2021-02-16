VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMailMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


'FOR INFO ON HOW PROCEDURE WORKS, GO HERE:
    'https://rtmccormick.com/2015/11/23/automatic-mail-merge-with-vba-and-access/

Private Sub cmdStartMerge_Click()
    Dim msg, Style, Response
    msg = "Mail Merge Successful, would you like to view your completed files?"
    Style = vbYesNo
    
    DoCmd.runSQL ("DELETE * FROM tblMaileMerge")
    
    DoCmd.OpenQuery ("qappMailMerge")
    
    MsgBox Me.fraOutput.Value
    
    'check if regular letters exist and need printed (i.e. hics letters w/o the appeal form)
    If DCount("name", "tblMaileMerge", "appeal_form = 0") > 0 Then
        Call startMerge(Me.fraOutput.Value)
    End If
    
    'pause to ensure all database processes are finished before running next mail merge
    DBEngine.Idle dbRefreshCache
    
    'check if letters  w/ appeal forms need printed
    If DCount("name", "tblMaileMerge", "appeal_form = -1") > 0 Then
        Call startMergeAppealForm(Me.fraOutput.Value)
    End If
    
    Response = MsgBox(msg, Style)
    If Response = vbYes Then
        Me.PrintComplete.Enabled = True
       Shell "explorer.exe " & "X:\Member Enrollment\Member Enrollment(Custom)\Marketplace\HICS\Resolution Letters and Report\db\", vbNormalFocus
       End
    End If
    
    'use 'END' statement to reset all module-level variables and all static local variables in all modules.
    'see this webpage for more info: https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/end-statement
    End

End Sub

'------------------------------------------------------------
' PrintComplete_Click
'------------------------------------------------------------
Private Sub PrintComplete_Click()
On Error GoTo PrintComplete_Click_Err
Dim usrinput
usrinput = MsgBox("Have all letters printed Successfully?", vbYesNo, "Succesful Print?")
    If usrinput = vbYes Then
    DoCmd.OpenQuery "qupdPrinted", acViewNormal, acEdit
    ElseIf usrinput = vbNo Then
        Exit Sub
        MsgBox ("You must complete printing before completing the mailing process")
    End If

PrintComplete_Click_Exit:
    Exit Sub

PrintComplete_Click_Err:
    MsgBox Error$
    Resume PrintComplete_Click_Exit

End Sub

'
